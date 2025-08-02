//go:build !headless
// +build !headless

package main

import "github.com/bytesyntax/schedulehelper/internal/gui"

func main() {
	gui.RunStandalone()
}
