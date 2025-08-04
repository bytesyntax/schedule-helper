.PHONY: all desktop mobile linux darwin windows android ios

app-name = github.com/bytesyntax/schedulehelper
app-version = 1.0.0

linux_arch   = amd64,arm64
windows_arch = amd64
darwin_arch  = amd64
android_arch = arm64
ios_arch     =

all: desktop mobile

desktop: linux windows darwin

mobile: android ios

linux windows darwin android ios:
	fyne-cross $@ --arch=$(${@}_arch) --app-id=$(app-name) --app-version=$(app-version) --name=schedulehelper ./cmd/schedulehelper

