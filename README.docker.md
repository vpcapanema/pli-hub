# Deploy do PLI Hub com Docker (sem Render)

Este projeto e estatico e pode ser executado como container Nginx.

## 1) Build e subida local/VM

```bash
docker compose up -d --build
```

A aplicacao ficara disponivel em `http://SEU_HOST:8080`.

## 2) Atualizar depois de mudanças

```bash
docker compose down
docker compose up -d --build
```

## 3) Logs

```bash
docker compose logs -f
```

## 4) Publicar em porta 80 (opcional)

Edite `docker-compose.yml` e troque:

```yaml
ports:
  - "8080:80"
```

por:

```yaml
ports:
  - "80:80"
```

## 5) Executar em boot

Ja esta configurado com `restart: unless-stopped`.
