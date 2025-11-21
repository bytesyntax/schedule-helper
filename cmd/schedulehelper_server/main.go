//go:build headless
// +build headless

package main

import "github.com/bytesyntax/schedule-helper/internal/core"

func main() {
	core.RunHeadless()
}
