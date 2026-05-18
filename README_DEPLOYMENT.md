# Deploy em nuvem

Este repositório contém um app Flask para processar PDFs e gerar Excel.

## O que já foi preparado

- `Dockerfile` com Python 3.14 e instalação de `poppler-utils`
- `requirements.txt` com `flask` e `openpyxl`

## Deploy recomendado (gratuito)

### Render.com

1. Crie uma conta em https://render.com
2. Conecte seu repositório GitHub que contenha este projeto
3. Crie um novo serviço Web:
   - Tipo: `Web Service`
   - Build Command: `docker build -t app .`
   - Start Command: `python "Leitor de PDF.py"`
4. Render usará a porta `5000` exposta pelo `Dockerfile`.
5. Após o deploy, use a URL pública fornecida pelo Render.

### Railway.app

1. Crie uma conta em https://railway.app
2. Conecte seu repositório GitHub
3. Crie um novo projeto e adicione um serviço Docker
4. O serviço roda a partir do `Dockerfile` existente.
5. A Railway fornecerá uma URL pública para o app.

## Observações

- A app já usa `host="0.0.0.0"` e `port=5000`, então está pronta para receber conexões externas.
- O PDF só será processado corretamente se o `pdftotext` estiver disponível, o que o `Dockerfile` já resolve.
