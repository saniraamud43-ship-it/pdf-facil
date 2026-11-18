# PDF Fácil

Aplicativo web progressivo (PWA) para ferramentas PDF, com interface responsiva em português.

## Executar
- Abra `index.html` em um navegador moderno para testar as funções básicas.
- Para instalar como PWA, publique a pasta em uma hospedagem HTTPS estática (por exemplo, GitHub Pages ou Netlify) e abra o endereço no celular/computador.
- O primeiro carregamento exige internet para baixar PDF-Lib e PDF.js via CDN. Depois, a interface básica pode ser armazenada pelo service worker; as bibliotecas CDN ainda precisam estar disponíveis.

## Recursos implementados
- Mesclar PDFs
- Extrair páginas por intervalos
- PDF para PNG (cada página gera um download)
- Imagens JPG/PNG para PDF
- Girar todas as páginas
- Regravar/otimizar PDF (não garante redução de tamanho)

## Limitações conhecidas
- PDF para Word/DOCX ainda não implementado. A conversão fiel exige reconstruir layout, tabelas e imagens; extração simples de texto não preserva a diagramação.
- Não há OCR para PDFs digitalizados.
- Arquivos grandes podem consumir muita memória no celular.
- A conversão PDF para PNG pode iniciar vários downloads, dependendo do navegador.
- Todo processamento implementado ocorre no navegador. Não existe backend nem upload de arquivos.

## Tecnologias
HTML, CSS, JavaScript, PDF-Lib e PDF.js.
