.PHONY: install run-server run-client lint format

install:
	pip install -r requirements.txt

run-server:
	python surgibot_server.py

run-client:
	python surgibot_client.py

lint:
	ruff check src

format:
	black src
