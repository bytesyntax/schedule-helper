//go:build headless
// +build headless

package main

import "github.com/bytesyntax/schedulehelper/internal/core"

func main() {
	core.RunHeadless()
}
