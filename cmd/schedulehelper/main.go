//go:build !headless
// +build !headless

package main

import "github.com/bytesyntax/schedule-helper/internal/gui"

func main() {
	gui.RunStandalone()
}
