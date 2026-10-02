package staticref

import (
	"fmt"
	"strings"
	"unicode"
	"unicode/utf8"
)

// checkFunctions lexes only function calls for the profile's smaller allowlist.
// The production formula parser subsequently validates the entire expression.
var allowedFunctions = map[string][2]int{
	"IF": {2, 3}, "SUM": {1, 255}, "MIN": {1, 255}, "MAX": {1, 255}, "AVERAGE": {1, 255}, "COUNT": {1, 255}, "COUNTA": {1, 255}, "AND": {1, 255}, "OR": {1, 255}, "LOG10": {1, 1}, "ABS": {1, 1}, "NOT": {1, 1}, "ROUND": {2, 2},
}

func checkFunctions(source string) error {
	type frame struct {
		name     string
		args     int
		hasToken bool
	}
	var stack []frame
	for at := 0; at < len(source); {
		r, n := utf8.DecodeRuneInString(source[at:])
		if forbiddenControl(r) || r != ' ' && unicode.IsSpace(r) {
			return refuse("unsupported-static-reference", fmt.Errorf("unsupported formula character"))
		}
		if r == ' ' {
			at += n
			continue
		}
		if r == '"' || r == '\'' {
			q := r
			at += n
			closed := false
			for at < len(source) {
				v, w := utf8.DecodeRuneInString(source[at:])
				if forbiddenControl(v) {
					return refuse("unsupported-static-reference", fmt.Errorf("unsupported formula character"))
				}
				at += w
				if v == q {
					if at < len(source) {
						next, z := utf8.DecodeRuneInString(source[at:])
						if next == q {
							at += z
							continue
						}
					}
					closed = true
					break
				}
			}
			if !closed {
				return refuse("unsupported-static-reference", fmt.Errorf("unclosed quote"))
			}
			if len(stack) > 0 {
				stack[len(stack)-1].hasToken = true
			}
			continue
		}
		if unicode.IsLetter(r) || r == '_' || r == '$' {
			start := at
			at += n
			for at < len(source) {
				v, w := utf8.DecodeRuneInString(source[at:])
				if !unicode.IsLetter(v) && !unicode.IsDigit(v) && v != '_' && v != '.' && v != '$' {
					break
				}
				at += w
			}
			word := source[start:at]
			look := at
			for look < len(source) {
				v, w := utf8.DecodeRuneInString(source[look:])
				if v != ' ' {
					break
				}
				look += w
			}
			if look < len(source) && source[look] == '(' {
				name := strings.ToUpper(word)
				if _, ok := allowedFunctions[name]; !ok {
					return refuse("unsupported-static-reference", fmt.Errorf("unsupported function %s", word))
				}
			}
			if len(stack) > 0 {
				stack[len(stack)-1].hasToken = true
			}
			continue
		}
		switch r {
		case '(':
			// Find a preceding function identifier without interpreting strings.
			end := at
			start := end
			for start > 0 {
				v, w := utf8.DecodeLastRuneInString(source[:start])
				if unicode.IsLetter(v) || unicode.IsDigit(v) || v == '_' || v == '.' {
					start -= w
				} else {
					break
				}
			}
			name := ""
			if start < end {
				name = strings.ToUpper(source[start:end])
			}
			stack = append(stack, frame{name: name})
			if len(stack) > 129 {
				return refuse("static-reference-limit", fmt.Errorf("formula depth limit"))
			}
		case ',':
			if len(stack) > 0 {
				top := &stack[len(stack)-1]
				if !top.hasToken {
					return refuse("unsupported-static-reference", fmt.Errorf("blank function argument"))
				}
				top.args++
				top.hasToken = false
			}
		case ')':
			if len(stack) > 0 {
				top := stack[len(stack)-1]
				stack = stack[:len(stack)-1]
				if top.name != "" {
					arity, ok := allowedFunctions[top.name]
					if !ok {
						return refuse("unsupported-static-reference", fmt.Errorf("unsupported function %s", top.name))
					}
					if top.args > 0 && !top.hasToken {
						return refuse("unsupported-static-reference", fmt.Errorf("blank function argument"))
					}
					if top.hasToken {
						top.args++
					}
					if top.args < arity[0] || top.args > arity[1] {
						return refuse("unsupported-static-reference", fmt.Errorf("unsupported function arity"))
					}
				}
				if len(stack) > 0 {
					stack[len(stack)-1].hasToken = true
				}
			}
		default:
			if len(stack) > 0 && r != ',' {
				stack[len(stack)-1].hasToken = true
			}
		}
		at += n
	}
	return nil
}

func forbiddenControl(r rune) bool { return r < 0x20 || r >= 0x7f && r <= 0x9f }
