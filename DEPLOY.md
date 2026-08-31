# Como versionar e publicar o PLI Hub

## Visao geral

O PLI Hub e um site estatico. O conteudo vive neste repositorio
(`github.com/vpcapanema/pli-hub`, branch `master`) e e servido em producao por um
container Nginx na instancia EC2 `56.125.163.194`.

Caminho de uma requisicao em producao:

```
https://56.125.163.194/pli-hub/
  -> nginx do host (Ubuntu, /etc/nginx/snippets/pli-hub-subpath.conf)
  -> proxy_pass http://127.0.0.1:8080/pli-hub/
  -> container "pli-hub" (nginx:1.27-alpine, nginx/default.conf)
  -> arquivos estaticos copiados na imagem
```

Como os arquivos sao copiados para dentro da imagem no `docker build`, **editar
arquivo no servidor nao muda nada**: e preciso reconstruir o container.

## 1) Commit (maquina local)

```bash
cd d:/REPOSITORIOS/pli-hub
git add -A
git commit -m "feat: descricao do que mudou"
git push origin master
```

Convencao de mensagem usada no historico: `feat:`, `fix:`, `refactor:` seguido de
descricao curta em portugues.

Se a mudanca altera a lista de aplicacoes ou qualquer arquivo do PRECACHE, suba
tambem a versao do cache do service worker, senao os navegadores continuam
servindo a versao antiga:

- `sw.js`: `CACHE_VERSION = "pli-hub-vN"` -> `pli-hub-vN+1`
- `index.html`: `sw.js?v=AAAAMMDD` -> data do dia

## 2) Deploy (servidor)

```bash
ssh ubuntu@56.125.163.194
cd /home/ubuntu/local-github-vm
./deploy/deploy.sh
```

O `deploy/deploy.sh` faz `git pull --ff-only origin master`, reconstroi a imagem,
sobe o container e confere `/health` e `/pli-hub/`.

No Windows, sem cliente SSH configurado, a mesma coisa com o PuTTY e a chave
`SRV-SISTEMA-30001480.ppk`:

```bash
plink -ssh -i CAMINHO\SRV-SISTEMA-30001480.ppk ubuntu@56.125.163.194 \
  "cd /home/ubuntu/local-github-vm && ./deploy/deploy.sh"
```

## 3) Conferir

```bash
curl -sk https://56.125.163.194/pli-hub/ | grep -o "name:'[^']*'"
curl -sk https://56.125.163.194/pli-hub/sw.js | grep CACHE_VERSION
```

No navegador, recarregue uma vez com Ctrl+Shift+R.

## Notas

- O diretorio de producao chama-se `local-github-vm` por motivo historico; e um
  clone deste repositorio. Nao renomeie sem antes derrubar o container, porque o
  nome do projeto do Docker Compose vem do nome do diretorio.
- Nao ha CI: `git push` sozinho **nao** publica nada. O passo 2 e obrigatorio.
- Nao edite arquivos direto no servidor. Qualquer alteracao local ali e perdida
  no proximo `git pull --ff-only`.

## Ambiente local para testar antes de publicar

```bash
docker compose up -d --build   # http://localhost:8080/pli-hub/
docker compose down
```
