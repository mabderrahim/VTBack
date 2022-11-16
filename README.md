# Visite technique - Backend

This program is the backend of the technical visit application.

## Getting Started

To use this program you can use Docker.

### With Docker

The following steps will guide you through the program set up with Docker.

#### Prerequisites

Docker is required. Information about installation are available [here](https://docs.docker.com/desktop/)

#### Install

Build a docker image called `technical_visit_backend`.

```
docker build --tag technical_visit_backend .
```

#### Run

Instantiate a container called `technical_visit_backend`. Replace `5001` with the port on which the backend should be reachable.

```
docker run -d -p 5001:5000 --name technical_visit_backend technical_visit_backend
```

#### Uninstall

Uninstall the server with the following commands

```
docker rm -f technical_visit_backend
docker image rm technical_visit_backend
```

### Without Docker

The following steps will guide you through the program set up without Docker.

#### Prerequisites

Python 3.7 (or higher) is required. You can download it from [here](https://www.python.org/downloads/).

#### Install

Install the required libraries with the following command :

```
pip install -r requirements.txt
```

#### Run

Run this program with the following command :

```
python app.py
```

## License

This program is proprietary.
<br>See [LICENSE](LICENSE) for more information.