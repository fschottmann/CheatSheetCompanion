---
aliases:
  - docker compose
  - container
tags:
  - dev
---
# Docker

## Containers

```bash
docker ps -a
docker run -d --name web -p 8080:80 nginx
docker logs -f web
docker exec -it web sh
docker stop web && docker rm web
```

## Images and cleanup

```bash
docker images
docker build -t myapp:1.0 .
docker image prune -a
docker system df
docker system prune --volumes
```

## Compose

```bash
docker compose up -d
docker compose logs -f api
docker compose pull && docker compose up -d
docker compose down -v
```

```yaml
services:
  web:
    image: nginx:alpine
    ports:
      - "8080:80"
    volumes:
      - ./html:/usr/share/nginx/html:ro
```

See [[Linux Shell]] for the host side.
