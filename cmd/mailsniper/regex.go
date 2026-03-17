package cmd

import "regexp"

// compileRegex compiles a regex string and returns it.
func compileRegex(pattern string) (*regexp.Regexp, error) {
	return regexp.Compile(pattern)
}
