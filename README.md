# Gerador de Folhas de Ponto para Professores

## Descrição
Este projeto automatiza a geração de folhas de ponto para professores com base em uma planilha de controle de presença. A partir do arquivo `Livro.xlsx`, o script gera um documento `.docx` para cada professor, preenchendo automaticamente os campos de nome, turno, situação, cadastro, vínculo e horários das aulas de acordo com o modelo `Março.docx`.

## Funcionalidades
- **Leitura da planilha de presença** (`.xlsx`)
- **Geração automática de documentos** no formato `.docx`
- **Preenchimento dinâmico** dos dados dos professores
- **Criação de pastas** para organizar os arquivos gerados
- **Correção de nomes de arquivos** para evitar caracteres inválidos

## Requisitos
Antes de executar o projeto, instale as dependências necessárias:

```bash
pip install pandas openpyxl xlrd python-docx
```

## Como Usar
1. **Coloque os arquivos na pasta do projeto:**
   - `LIVRO.xlsx`
   - `Março.docx`
2. **Execute o script:**
   ```bash
   python main.py
   ```
3. **Verifique a pasta de saída:**
   - Os arquivos gerados estarão em `folhas_ponto/`

## Estrutura do Projeto
```
projeto_mae/
│-- Livro.xlsx
│-- Março.docx
│-- main.py
│-- folhas_ponto/  
│-- README.md
```


