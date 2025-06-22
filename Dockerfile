FROM golang:1.24.4-bookworm AS builder

WORKDIR /build

COPY ./src/* ./

RUN go mod download && go mod verify

RUN CGO_ENABLED=0 GOOS=linux go build -o app .

FROM scratch

COPY --from=builder /build/app /build/upload.html /

CMD [ "/app" ]