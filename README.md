# Clockify Report Converter

Aplicativo desktop para converter relatórios exportados do Clockify em planilhas Excel formatadas e profissionais.

## O que o aplicativo faz?

O **Clockify Report Converter** transforma o arquivo de exportação Detailed do Clockify em um relatório Excel formatado com:

- **Aba Summary Report**: Resumo por projeto com totais de horas e valores (gerado automaticamente)
- **Aba Detailed Report**: Registro detalhado de todas as entradas de tempo

O relatório gerado inclui:

- Formatação profissional com cores e estilos
- Cálculo automático de valores baseado na taxa horária configurada
- Nome do usuário no nome do arquivo para fácil identificação
- Totais calculados automaticamente com precisão total (sem erros de arredondamento)

## Como executar o aplicativo

### Opção 1: Executável Windows (Recomendado para usuários finais)

1. Acesse a página de [Releases](https://github.com/mauricio-jmmendes/ClockifyReportManager/releases) do repositório
2. Baixe o arquivo `Clockify.Report.Converter.exe` da versão mais recente
3. Coloque o executável na mesma pasta dos arquivos do Clockify
4. Dê duplo clique para abrir

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

O executável será criado na pasta `dist/`

## Como usar

### Passo 1: Exportar relatório do Clockify

No Clockify, exporte o **Detailed Report** em formato Excel (.xlsx):

- **Detailed Report** - Relatório detalhado com todas as entradas
  - Nome esperado: `Clockify_Time_Report_Detailed_DD_MM_YYYY-DD_MM_YYYY.xlsx`

> **Nota:** O Summary Report é gerado automaticamente a partir do Detailed Report, não é necessário exportá-lo separadamente.

### Passo 2: Preparar o arquivo

Coloque o arquivo exportado na mesma pasta do aplicativo. O app detectará automaticamente o arquivo ao iniciar.

### Passo 3: Configurar e converter

1. **Abra o aplicativo**

2. **Verifique o arquivo de entrada**
   - O campo "Clockify Detailed Report" deve mostrar o arquivo detectado
   - Se não foi detectado, clique em "Browse" para selecionar manualmente

3. **Preencha as configurações:**
   - **User Name**: Seu nome (será usado no nome do arquivo de saída)
   - **Billable Rate (BRL/hour)**: Taxa horária em Reais (padrão: 50)
   - **Output Folder**: Pasta onde o arquivo será salvo

4. **Clique em "Convert Reports"**

5. **Arquivo de saída**
   - O arquivo será salvo com o nome: `Nome_Usuario_Time_Report_DD_MM_YYYY-DD_MM_YYYY.xlsx`
   - Se já existir um arquivo com o mesmo nome, você pode escolher sobrescrever ou criar um novo com numeração

## Modo CLI (Linha de Comando)

Para usuários avançados ou automação, existe também uma versão CLI do conversor.

### Uso básico

```bash
# Auto-detectar arquivo Detailed na pasta atual (taxa padrão: 50)
python clockify_report_converter.py

# Com taxa personalizada
python clockify_report_converter.py --rate 150

# Especificando arquivo manualmente
python clockify_report_converter.py --rate 150 \
    --detailed Clockify_Time_Report_Detailed_01_12_2025-26_12_2025.xlsx \
    --output Relatorio_Saida.xlsx
```

### Opções disponíveis

| Opção | Descrição |
|-------|-----------|
| `--rate <valor>` | Taxa horária em BRL (padrão: 50) |
| `--detailed <arquivo>` | Caminho para o arquivo Detailed |
| `--output <arquivo>` | Caminho para o arquivo de saída |

### Quando usar o modo CLI?

- Automação e scripts em lote
- Integração com outros sistemas
- Ambientes sem interface gráfica (servidores)
- Usuários que preferem linha de comando

## Estrutura do arquivo gerado

### Aba "Summary Report"

| Coluna | Descrição |
|--------|-----------|
| Project | Nome do projeto (com cliente) |
| Description | Descrição das tarefas |
| Time (h) | Tempo em formato hh:mm:ss |
| Time (decimal) | Tempo em formato decimal |
| Amount (BRL) | Valor calculado (tempo × taxa) |

> O Summary é gerado automaticamente agrupando os dados do Detailed Report por projeto e descrição.

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
| Duration (h) | Duração em formato hh:mm:ss |
| Duration (decimal) | Duração em formato decimal |
| Billable Rate | Taxa horária |
| Billable Amount | Valor calculado |

## Precisão dos cálculos

O aplicativo utiliza a coluna `Duration (h)` (formato HH:MM:SS) para todos os cálculos de tempo, garantindo precisão total. Isso evita os erros de arredondamento que ocorrem quando se usa a coluna `Duration (decimal)` pré-arredondada do Clockify.

## Resolução de problemas

### "Permission denied" ao salvar

- O arquivo de destino pode estar aberto no Excel. Feche-o e tente novamente.

### Arquivo não detectado automaticamente

- Verifique se o nome do arquivo segue o padrão do Clockify: `Clockify_Time_Report_Detailed_*.xlsx`
- Use o botão "Browse" para selecionar manualmente

### Erro ao ler o arquivo

- Certifique-se de que o arquivo é uma exportação válida do Clockify em formato .xlsx

### Totais não coincidem com o Clockify

- Os totais do aplicativo são calculados com precisão total a partir do Duration (h)
- O Clockify pode mostrar valores ligeiramente diferentes devido ao arredondamento

## Dependências

- Python 3.10+
- customtkinter >= 5.2.0
- pandas >= 2.0.0
- openpyxl >= 3.1.0
- pyinstaller >= 6.0.0 (apenas para build)

## Licença

Este projeto é de uso livre.
