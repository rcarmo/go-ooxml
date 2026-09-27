// Package formula recognises a conservative static A1 expression subset. It is
// dependency analysis, never a calculator; unsupported syntax must refuse.
package formula

import (
	"fmt"
	"strconv"
	"strings"
	"unicode"
	"unicode/utf8"
)

type Cell struct {
	Row, Column                 int
	AbsoluteRow, AbsoluteColumn bool
}
type Reference struct {
	Sheet       string
	First, Last Cell
	Start, End  int
}
type token struct {
	kind, text string
	start, end int
}
type parser struct {
	tokens    []token
	at, depth int
	refs      []Reference
}

// Analyze returns static scalar/range references and exact UTF-8 byte offsets.
// Strings, booleans, arithmetic/comparison and a closed set of nonvolatile
// functions are accepted. Names, UDFs, external/3D/structured/dynamic references,
// arrays, spills, intersections and unknown syntax refuse with no partial result.
func Analyze(source string) ([]Reference, error) {
	if !utf8.ValidString(source) || len(source) > 1<<20 {
		return nil, fmt.Errorf("invalid or excessive formula text")
	}
	tokens, err := lex(source)
	if err != nil {
		return nil, err
	}
	p := parser{tokens: tokens, refs: []Reference{}}
	if p.peek().kind == "punct" && p.peek().text == "=" {
		p.at++
	}
	if err = p.expression(0); err != nil {
		return nil, err
	}
	if p.peek().kind != "end" {
		return nil, fmt.Errorf("unsupported trailing token at byte%d", p.peek().start)
	}
	return p.refs, nil
}
func lex(s string) ([]token, error) {
	out := []token{}
	for at := 0; at < len(s); {
		r, n := utf8.DecodeRuneInString(s[at:])
		if r == ' ' || r == '\t' || r == '\r' || r == '\n' {
			at += n
			continue
		}
		start := at
		if r == '"' || r == '\'' {
			quote := byte(r)
			at++
			var b strings.Builder
			closed := false
			for at < len(s) {
				if s[at] == quote {
					if at+1 < len(s) && s[at+1] == quote {
						b.WriteByte(quote)
						at += 2
						continue
					}
					at++
					closed = true
					break
				}
				rr, nn := utf8.DecodeRuneInString(s[at:])
				b.WriteRune(rr)
				at += nn
			}
			if !closed {
				return nil, fmt.Errorf("unterminated formula string/sheet")
			}
			kind := "string"
			if quote == '\'' {
				kind = "sheet"
			}
			out = append(out, token{kind, b.String(), start, at})
		} else if r >= '0' && r <= '9' || r == '.' && at+1 < len(s) && s[at+1] >= '0' && s[at+1] <= '9' {
			for at < len(s) && s[at] >= '0' && s[at] <= '9' {
				at++
			}
			if at < len(s) && s[at] == '.' {
				at++
				for at < len(s) && s[at] >= '0' && s[at] <= '9' {
					at++
				}
			}
			if at < len(s) && (s[at] == 'e' || s[at] == 'E') {
				at++
				if at < len(s) && (s[at] == '+' || s[at] == '-') {
					at++
				}
				digits := at
				for at < len(s) && s[at] >= '0' && s[at] <= '9' {
					at++
				}
				if at == digits {
					return nil, fmt.Errorf("invalid exponent")
				}
			}
			out = append(out, token{"number", s[start:at], start, at})
		} else if unicode.IsLetter(r) || r == '_' || r == '$' {
			at += n
			for at < len(s) {
				rr, nn := utf8.DecodeRuneInString(s[at:])
				if !unicode.IsLetter(rr) && !unicode.IsDigit(rr) && rr != '_' && rr != '.' && rr != '$' {
					break
				}
				at += nn
			}
			out = append(out, token{"word", s[start:at], start, at})
		} else {
			if !strings.ContainsRune("+-*/^&%=<>(),:!", r) {
				return nil, fmt.Errorf("unsupported token %q at byte%d", r, at)
			}
			at += n
			if (r == '<' || r == '>') && at < len(s) && (s[at] == '=' || r == '<' && s[at] == '>') {
				at++
			}
			out = append(out, token{"punct", s[start:at], start, at})
		}
		if len(out) > 100000 {
			return nil, fmt.Errorf("formula token budget exceeded")
		}
	}
	out = append(out, token{"end", "", len(s), len(s)})
	return out, nil
}
func (p *parser) peek() token { return p.tokens[p.at] }
func (p *parser) take(text string) bool {
	if p.peek().kind == "punct" && p.peek().text == text {
		p.at++
		return true
	}
	return false
}
func precedence(op string) int {
	switch op {
	case "=", "<>", "<", ">", "<=", ">=":
		return 1
	case "&":
		return 2
	case "+", "-":
		return 3
	case "*", "/":
		return 4
	case "^":
		return 5
	}
	return -1
}
func (p *parser) expression(min int) error {
	p.depth++
	defer func() { p.depth-- }()
	if p.depth > 128 {
		return fmt.Errorf("formula depth limit")
	}
	if err := p.primary(); err != nil {
		return err
	}
	for {
		if p.take("%") {
			continue
		}
		if p.peek().kind != "punct" {
			return nil
		}
		prec := precedence(p.peek().text)
		if prec < min {
			return nil
		}
		p.at++
		if err := p.expression(prec + 1); err != nil {
			return err
		}
	}
}
func (p *parser) primary() error {
	t := p.peek()
	if t.kind == "end" {
		return fmt.Errorf("missing formula operand")
	}
	if p.take("+") || p.take("-") {
		return p.expression(6)
	}
	if p.take("(") {
		if err := p.expression(0); err != nil {
			return err
		}
		if !p.take(")") {
			return fmt.Errorf("unclosed formula parentheses")
		}
		return nil
	}
	if t.kind == "number" || t.kind == "string" {
		p.at++
		return nil
	}
	if t.kind != "word" && t.kind != "sheet" {
		return fmt.Errorf("unexpected token %q", t.text)
	}
	p.at++
	if p.take("!") {
		if t.text == "" || strings.ContainsAny(t.text, "[]:*?/\\") {
			return fmt.Errorf("invalid or external sheet reference")
		}
		first := p.peek()
		if first.kind != "word" {
			return fmt.Errorf("expected cell after sheet")
		}
		p.at++
		return p.reference(t.text, t.start, first)
	}
	if t.kind == "sheet" {
		return fmt.Errorf("quoted sheet lacks cell reference")
	}
	if p.take("(") {
		name := strings.ToUpper(t.text)
		arity, ok := functions[name]
		if !ok {
			return fmt.Errorf("unsupported or dynamic function %s", name)
		}
		count := 0
		if !p.take(")") {
			for {
				if err := p.expression(0); err != nil {
					return err
				}
				count++
				if p.take(")") {
					break
				}
				if !p.take(",") {
					return fmt.Errorf("function separator expected")
				}
			}
		}
		if count < arity[0] || count > arity[1] {
			return fmt.Errorf("unsupported arity for %s", name)
		}
		return nil
	}
	if strings.EqualFold(t.text, "TRUE") || strings.EqualFold(t.text, "FALSE") {
		return nil
	}
	return p.reference("", t.start, t)
}

var functions = map[string][2]int{"SUM": {1, 255}, "AVERAGE": {1, 255}, "MIN": {1, 255}, "MAX": {1, 255}, "COUNT": {1, 255}, "COUNTA": {1, 255}, "PRODUCT": {1, 255}, "IF": {2, 3}, "IFERROR": {2, 2}, "AND": {1, 255}, "OR": {1, 255}, "NOT": {1, 1}, "ABS": {1, 1}, "ROUND": {2, 2}, "ROUNDUP": {2, 2}, "ROUNDDOWN": {2, 2}, "INT": {1, 1}, "MOD": {2, 2}, "POWER": {2, 2}, "SQRT": {1, 1}, "LOG10": {1, 1}}

func (p *parser) reference(sheet string, start int, first token) error {
	cell, err := ParseCell(first.text)
	if err != nil {
		return err
	}
	last := cell
	end := first.end
	if p.take(":") {
		t := p.peek()
		if t.kind != "word" {
			return fmt.Errorf("range endpoint expected")
		}
		p.at++
		last, err = ParseCell(t.text)
		if err != nil {
			return err
		}
		end = t.end
	}
	p.refs = append(p.refs, Reference{sheet, cell, last, start, end})
	return nil
}

// ParseCell preserves absolute markers and enforces the XLSX grid bounds.
func ParseCell(text string) (Cell, error) {
	out := Cell{}
	at := 0
	if strings.HasPrefix(text, "$") {
		out.AbsoluteColumn = true
		at++
	}
	begin := at
	for at < len(text) {
		c := text[at]
		if c >= 'a' && c <= 'z' {
			c -= 32
		}
		if c < 'A' || c > 'Z' {
			break
		}
		out.Column = out.Column*26 + int(c-'A'+1)
		at++
		if at-begin > 3 {
			return Cell{}, fmt.Errorf("unsupported name/cell %q", text)
		}
	}
	if at == begin || out.Column > 16384 {
		return Cell{}, fmt.Errorf("invalid column %q", text)
	}
	if at < len(text) && text[at] == '$' {
		out.AbsoluteRow = true
		at++
	}
	begin = at
	for at < len(text) && text[at] >= '0' && text[at] <= '9' {
		at++
	}
	if at != len(text) || at == begin || text[begin] == '0' {
		return Cell{}, fmt.Errorf("unsupported name/cell %q", text)
	}
	row, err := strconv.ParseUint(text[begin:], 10, 32)
	if err != nil || row < 1 || row > 1048576 {
		return Cell{}, fmt.Errorf("invalid row %q", text)
	}
	out.Row = int(row)
	return out, nil
}
