# Clockify Report Converter

Aplicativo desktop para converter relatórios exportados do Clockify em planilhas Excel formatadas e profissionais.

## O que o aplicativo faz?

O **Clockify Report Converter** transforma os arquivos de exportação do Clockify (Summary e Detailed) em um relatório Excel formatado com:

- **Aba Summary Report**: Resumo por projeto com totais de horas e valores
- **Aba Detailed Report**: Registro detalhado de todas as entradas de tempo

O relatório gerado inclui:

- Formatação profissional com cores e estilos
- Cálculo automático de valores baseado na taxa horária configurada
- Nome do usuário no nome do arquivo para fácil identificação
- Totais calculados automaticamente

## Como executar o aplicativo

### Opção 1: Executável Windows (Recomendado para usuários finais)

1. Baixe o arquivo `Clockify Report Converter.exe` da pasta `dist/`
2. Coloque o executável na mesma pasta dos arquivos do Clockify
3. Dê duplo clique para abrir

**Não é necessário ter Python instalado!**

### Opção 2: Executar via Python (Para desenvolvedores)

**Pré-requisitos:**

- Python 3.10 ou superior

**Instalação:**

```bash
# Clone ou baixe o repositório
cd ClockifyReportManager

# Instale as dependências
pip install -r requirements.txt

# Execute o aplicativo
python clockify_app.py
```

### Opção 3: Gerar o executável você mesmo

```bash
# Instale as dependências
pip install -r requirements.txt

# Execute o script de build
build.bat
```

O executável será criado em `dist/Clockify Report Converter.exe`

## Como usar

### Passo 1: Exportar relatórios do Clockify

No Clockify, exporte dois relatórios em formato Excel (.xlsx):

1. **Summary Report** - Relatório resumido por projeto
   - Nome esperado: `Clockify_Time_Report_Summary_DD_MM_YYYY-DD_MM_YYYY.xlsx`

2. **Detailed Report** - Relatório detalhado com todas as entradas
   - Nome esperado: `Clockify_Time_Report_Detailed_DD_MM_YYYY-DD_MM_YYYY.xlsx`

### Passo 2: Preparar os arquivos

Coloque os dois arquivos exportados na mesma pasta do aplicativo. O app detectará automaticamente os arquivos ao iniciar.

### Passo 3: Configurar e converter

1. **Abra o aplicativo**

2. **Verifique os arquivos de entrada**
   - Os campos "Summary Report" e "Detailed Report" devem mostrar os arquivos detectados
   - Se não foram detectados, clique em "Browse" para selecionar manualmente

3. **Preencha as configurações:**
   - **User Name**: Seu nome (será usado no nome do arquivo de saída)
   - **Billable Rate (BRL/hour)**: Taxa horária em Reais (padrão: 50)
   - **Output Folder**: Pasta onde o arquivo será salvo

4. **Clique em "Convert Reports"**

5. **Arquivo de saída**
   - O arquivo será salvo com o nome: `Nome_Usuario_Time_Report_DD_MM_YYYY-DD_MM_YYYY.xlsx`
   - Se já existir um arquivo com o mesmo nome, você pode escolher sobrescrever ou criar um novo com numeração

## Estrutura do arquivo gerado

### Aba "Summary Report"

| Coluna | Descrição |
|--------|-----------|
| Project | Nome do projeto |
| Description | Descrição das tarefas |
| Time (h) | Tempo em formato hh:mm:ss |
| Time (decimal) | Tempo em formato decimal |
| Amount (BRL) | Valor calculado (tempo × taxa) |

### Aba "Detailed Report"

| Coluna | Descrição |
|--------|-----------|
| Project | Nome do projeto |
| Client | Nome do cliente |
| Description | Descrição da tarefa |
| User | Nome do usuário |
| Tags | Tags associadas |
| Start Date/Time | Data e hora de início |
| End Date/Time | Data e hora de término |
| Duration | Duração (h:mm:ss e decimal) |
| Billable Rate | Taxa horária |
| Billable Amount | Valor calculado |

## Resolução de problemas

### "Permission denied" ao salvar

- O arquivo de destino pode estar aberto no Excel. Feche-o e tente novamente.

### Arquivos não detectados automaticamente

- Verifique se os nomes dos arquivos seguem o padrão do Clockify
- Use o botão "Browse" para selecionar manualmente

### Erro ao ler os arquivos

- Certifique-se de que os arquivos são exportações válidas do Clockify em formato .xlsx

## Dependências

- Python 3.10+
- customtkinter >= 5.2.0
- pandas >= 2.0.0
- openpyxl >= 3.1.0
- pyinstaller >= 6.0.0 (apenas para build)

## Licença

Este projeto é de uso livre.
