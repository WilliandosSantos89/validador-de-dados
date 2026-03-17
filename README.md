# ✅ Validador de Dados de Materiais Educacionais

Script Python que valida a qualidade de dados operacionais antes da publicação,     aplica regras de negócio por campo, identifica inconsistências e gera relatório de qualidade automaticamente.

---

## 💡 Problema Resolvido

Dados incorretos em planilhas operacionais causam falhas silenciosas: materiais publicados na plataforma errada, responsáveis não identificados, prazos em formato inválido. O validador intercepta esses problemas antes que cheguem à operação.

---

## ✅ Funcionalidades

- Validação de 6 campos por registro: ID, Material, Plataforma, Responsável, Prazo e Status
- Regras por campo:
  - **ID** — padrão `MAT-000`
  - **Plataforma** — lista de valores permitidos
  - **Responsável** — cadastro de responsáveis válidos
  - **Prazo** — formato `DD/MM/YYYY`
  - **Status** — lista de valores permitidos
  - **Material** — campo obrigatório com tamanho mínimo
- Relatório Excel com 3 abas: Resumo, Erros e Válidos
- Taxa de qualidade calculada automaticamente
- Log de auditoria em JSON com detalhes por registro
- Mensagens de erro descritivas com sugestão de correção

---

## 🛠️ Tecnologias

| Tecnologia | Uso |
|------------|-----|
| Python 3.11 | Linguagem principal |
| openpyxl | Leitura e geração de planilhas Excel |

---

## 📁 Estrutura do Projeto

```
validador-de-dados/
│
├── validar.py                      # Script principal
├── materiais.xlsx                  # Planilha de entrada
├── validacao_YYYYMMDD.xlsx         # Relatório gerado automaticamente
├── log_validacao_YYYYMMDD.json     # Log de auditoria
├── .gitignore
└── README.md
```

---

## ▶️ Como Executar

**Pré-requisito:**

```bash
pip install openpyxl
```

**Execução:**

```bash
python validar.py
```

A planilha `materiais.xlsx` deve estar na mesma pasta do script.

---

## 📋 Exemplo de Saída

```
Lendo planilha: materiais.xlsx
Validando 12 registro(s)...

Validação concluída.
Total          : 12
Válidos        : 8
Com erro       : 4
Total de erros : 4
Taxa qualidade : 66.7%

Registros com problema:
  Linha 2 | 1       | ID — ID fora do padrão MAT-000
  Linha 4 | MAT-003 | Plataforma — use: LMS Moodle, Portal do Aluno, App Arco
  Linha 6 | MAT-005 | Responsável — Campo obrigatório vazio
  Linha 7 | MAT-006 | Status — use: Publicado, Pendente, Atrasado, Erro
```

---

## 🔗 Parte de uma Série

Este projeto faz parte de uma série de scripts de automação para operações educacionais:

1. [planilha-automatizada](https://github.com/WilliandosSantos89/planilha-automatizada) — controle e relatório de materiais
2. [upload-automatizado](https://github.com/WilliandosSantos89/upload-automatizado) — upload para Google Drive
3. **validador-de-dados** — validação de qualidade antes da publicação

---

## 🗺️ Próximos Passos

- [ ] Regras de validação configuráveis via arquivo externo
- [ ] Integração com Google Sheets em tempo real
- [ ] Envio de alertas por e-mail quando taxa de qualidade cair abaixo do limite
- [ ] Interface web para upload e validação sem linha de comando

---

## 👤 Autor

**Willian dos Santos**
Desenvolvedor em formação | ADS | Administração
[LinkedIn](https://www.linkedin.com/in/willian-dos-santos) • [GitHub](https://github.com/WilliandosSantos89)