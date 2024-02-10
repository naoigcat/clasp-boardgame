.PHONY: build
build:
	docker compose build

.PHONY: login
login:
	docker compose run --rm --entrypoint sh clasp -c $$'clasp login & \n read url \n curl $$url \n wait'

.PHONY: clone
clone:
	docker compose run --rm clasp clone $$SCRIPT_ID

.PHONY: pull
pull:
	docker compose run --rm clasp pull

.PHONY: push
push:
	docker compose run --rm clasp push
