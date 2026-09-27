package main

import (
	"strings"
	"testing"
)

func TestNativeDiscoveryBatch(t *testing.T) {
	source := `package sample
 import "testing"
 func helper(t *testing.T){}
 func Testable(){}
 func TestItems(t *testing.T){
   cases:=[]struct{name string; value int}{{"blank",0},{name:"number",value:1}}
   for _,c:=range cases{t.Run(c.name,func(t *testing.T){read(c.value)})}
   t.Run("literal",func(t *testing.T){write()})
 }
 func FuzzRead(f *testing.F){f.Add("a")}
 func BenchmarkRead(b *testing.B){read(1)}
 func ExampleRead(){read(1)}
 `
	f, ds, err := scan("native_test.go", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if f.HelpersOnly || len(f.Hash) != 64 || len(ds) != 4 {
		t.Fatalf("%+v %+v", f, ds)
	}
	d := ds[0]
	if d.ID != "native_test.go::TestItems" || d.Coverage != "unreviewed" || len(d.ScenarioIDs) != 0 || len(d.Subtests) != 2 || !d.Subtests[0].Dynamic || d.Subtests[1].Name != "literal" {
		t.Fatalf("discovery %+v", d)
	}
	if len(d.Groups) != 1 || d.Groups[0].Rows != 2 || strings.Join(d.Groups[0].Names, ",") != "blank,number" {
		t.Fatalf("groups %+v", d.Groups)
	}
	for i, k := range []string{"test", "fuzz", "benchmark", "example"} {
		if ds[i].Kind != k {
			t.Fatalf("kind %+v", ds[i])
		}
	}
	for _, tc := range []struct {
		name, source string
		bad, helpers bool
	}{
		{"helper only", "package x; func helper(){}", false, true},
		{"malformed", "package x; func TestOops(", true, false},
		{"method not test", "package x; type X int; func (x X) TestMethod(){}", false, true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			f, ds, err := scan("x_test.go", []byte(tc.source))
			if (err != nil) != tc.bad {
				t.Fatal(err)
			}
			if !tc.bad && (f.HelpersOnly != tc.helpers || len(ds) != 0) {
				t.Fatalf("%+v %+v", f, ds)
			}
		})
	}
}
