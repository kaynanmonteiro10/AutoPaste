# Sistema de Cópia de Fotos e Texto para PowerPoint

Este sistema foi desenvolvido em Python utilizando Tkinter para criar uma interface gráfica que automatiza a inserção de fotos e textos em apresentações do PowerPoint. O programa captura imagens da área de transferência, formata automaticamente as fotos, preenche campos de texto extraídos da área de transferência e permite limpar os dados inseridos.

## Funcionalidades

- Captura automática de imagens da área de transferência.
- Armazena as imagens em uma pasta específica.
- Insere fotos automaticamente no PowerPoint com formatação predefinida.
- Extrai e preenche automaticamente campos de texto extraídos da área de transferência.
- Remove fotos e textos de slides do PowerPoint.
- Insere a marcação "RUPTURA TOTAL" no slide atual.
- Interface interativa para conferência e ajustes.

## Requisitos

Antes de rodar o sistema, certifique-se de ter os seguintes requisitos instalados:

- Python 3.x
- Bibliotecas:
  - tkinter (padrão do Python)
  - Pillow
  - pywin32

Para instalar as dependências, execute:

```bash
pip install pillow pywin32
```

## Como Usar

1. Execute o script Python.
2. O programa ficará ativo e monitorando a área de transferência.
3. Copie uma imagem para a área de transferência e ela será automaticamente salva e exibida na interface.
4. Copie um texto estruturado (ex: "FORNECEDOR: XYZ") e os campos serão preenchidos automaticamente.
5. Clique no botão **"Colar Fotos e Textos no PowerPoint"** para inserir os dados no slide atual.
6. Utilize os botões para limpar textos, imagens ou inserir "RUPTURA TOTAL".

## Estrutura do Projeto

```
|— sistema.py  # Arquivo principal do programa
|— fotos_copiadas/  # Pasta onde as imagens copiadas são armazenadas temporariamente
```

## Contribuição

Sinta-se à vontade para contribuir com melhorias no código, relatando bugs ou sugerindo novas funcionalidades.

