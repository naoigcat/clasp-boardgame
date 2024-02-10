# GAS BoardGameGeek

## Build

Rebuild Docker image.

```sh
make build
```

## Login

Login to clasp by accessing the URL output to the terminal, and pasting the authentication code to the terminal.

```sh
make login
```

## Clone

Download the script with `{SCRIPT_ID}` that can be obtained by the URL and save it.

```sh
make clone SCRIPT_ID={SCRIPT_ID}
```

## Push

Upload the script.

```sh
make push
```

## Pull

Download the script.

```sh
make pull
```
