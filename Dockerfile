FROM golang:1.24.5-bookworm AS builder

WORKDIR /build

COPY . .

RUN go mod tidy

RUN CGO_ENABLED=0 GOOS=linux go build -tags headless -o server ./cmd/schedulehelper_server

FROM scratch

COPY --from=builder /build/server /build/assets/upload.html /

CMD [ "/server" ]
