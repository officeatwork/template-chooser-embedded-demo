# officeatwork Template Chooser Embedded

This repository demonstrates how to embed the officeatwork Template Chooser within business applications.

## Live demo

[Click here](https://template-chooser-embedded-demo.officeatwork.com/) to see a live demo.

## Running locally
To run locally you can use any local web server. As Template Chooser Embedded runs in an iframe locally, you have to run the local web server with SSL (https).
For that use OpenSSL to create a self signed certificate.

Example with http-server:

```sh
npx http-server -S -C certificate.crt -K certificate.key
```

## Usage

The function `handleEvent` in `tc-embedded.js` shows how to consume events from Template Chooser Embedded.
