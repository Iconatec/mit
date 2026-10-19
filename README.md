# Gerador de JSON para MIT — Excel VBA
Macro VBA para montar um arquivo JSON destinado ao MIT a partir de duas abas do Excel: `MIT_Cabecalho` e `MIT_Debitos`.
## Objetivo
Automatizar a montagem dos dados de período de apuração, informações iniciais e débitos em um arquivo `.json`.
A rotina permite gerar arquivos com ou sem movimento e organiza os débitos por grupo e tipo de lista.
**Atenção:** o código gera o arquivo localmente. Não realiza transmissão, autenticação ou validação junto ao sistema de destino. A compatibilidade com o leiaute exigido deve ser conferida antes da utilização.
## Estrutura da pasta de trabalho
| Aba | Finalidade |
|---|---|
| `MIT_Cabecalho` | Período de apuração, dados iniciais e responsável |
| `MIT_Debitos` | Relação dos débitos que serão incluídos no JSON |
As abas devem estar no mesmo arquivo que contém a macro, pois o código utiliza `ThisWorkbook`.
## Aba MIT_Cabecalho
Os dados são lidos exclusivamente da linha 2.
| Coluna | Campo | Tratamento no código |
|---|---|---|
| A | `MesApuracao` | Convertido para inteiro com `CLng` |
| B | `AnoApuracao` | Convertido para inteiro com `CLng` |
| C | `SemMovimento` | Convertido para booleano |
| D | `QualificacaoPj` | Inserido no JSON sem aspas |
| E | `TributacaoLucro` | Incluído condicionalmente, sem aspas |
| F | `VariacoesMonetarias` | Obrigatório no código quando há movimento |
| G | `RegimePisCofins` | Incluído quando há movimento e está preenchido |
| H | `CpfResponsavel` | Mantém somente os dígitos |
| I | DDD do responsável | Mantém somente os dígitos |
| J | Telefone do responsável | Mantém somente os dígitos |
| K | E-mail do responsável | Remove espaços nas extremidades e aplica escape de texto |
| L | UF do registro CRC | Converte para maiúsculas |
| M | Número do registro CRC | Converte para maiúsculas e aplica escape de texto |
| N | `BalancoLucroReal` | Convertido para booleano |
| O | CNPJ do contribuinte | Mantém somente os dígitos e exige 14 posições |
O CNPJ da coluna O é utilizado para validar o comprimento e montar o nome do arquivo. Ele não é incluído como campo no JSON pela versão apresentada.
### Valores booleanos
Os campos `SemMovimento` e `BalancoLucroReal` consideram verdadeiros os seguintes conteúdos:
- `TRUE`
- `VERDADEIRO`
- `1`
- `SIM`
A comparação ignora maiúsculas/minúsculas e espaços nas extremidades. Qualquer outro conteúdo é interpretado como falso.
### Regras dos dados iniciais
- `SemMovimento` e `QualificacaoPj` são sempre incluídos.
- `TributacaoLucro` é incluído somente quando há movimento, a qualificação é diferente de `"11"` e o campo está preenchido.
- `VariacoesMonetarias` é incluído somente quando há movimento e está preenchido.
- `RegimePisCofins` é incluído somente quando há movimento e está preenchido.
- `ResponsavelApuracao` e `CpfResponsavel` são sempre incluídos, mesmo que o CPF esteja vazio.
- `TelResponsavel` é incluído somente quando DDD e telefone estão preenchidos.
- `EmailResponsavel` é incluído quando preenchido.
- `RegistroCrc` é incluído somente quando UF e número do registro estão preenchidos.
## Aba MIT_Debitos
A linha 1 deve conter os cabeçalhos. Os registros são lidos a partir da linha 2 até a última linha preenchida na coluna A.
| Coluna | Campo | Tratamento no código |
|---|---|---|
| A | Grupo do débito | Define em qual grupo o registro será incluído |
| B | Tipo de lista | Define lista normal ou após evento |
| C | `IdDebito` | Inserido sem aspas |
| D | `IdEventoDebito` | Incluído somente na lista após evento, quando preenchido |
| E | `CodigoDebito` | Mantém somente os dígitos e é inserido como texto |
| F | `PaDebito` | Incluído quando preenchido, exceto para `RetPagamentoUnificado` |
| G | `AnoPostergado` | Incluído quando preenchido, sem aspas |
| H | `TrimPostergado` | Incluído quando preenchido, sem aspas |
| I | `AnoDebito` | Incluído quando preenchido, sem aspas |
| J | `CnpjEstabelecimento` | Mantém somente os dígitos |
| K | `CnpjIncorporacao` | Mantém somente os dígitos |
| L | `CnpjScp` | Mantém somente os dígitos |
| M | `CodigoMunicipioOuro` | Mantém somente os dígitos |
| N | `ValorDebito` | Convertido com `CDbl` e exportado com duas casas decimais |
Os campos opcionais são omitidos quando estão vazios. O código não verifica se eles são obrigatórios em situações específicas do leiaute.
### Grupos reconhecidos
A coluna A deve utilizar um dos seguintes grupos:
```text
Irpj
Csll
Irrf
Ipi
Iof
PisPasep
Cofins
ContribuicoesDiversas
Cpss
RetPagamentoUnificado
```
A comparação ignora maiúsculas/minúsculas e espaços nas extremidades.
Grupos não reconhecidos não são exportados. A macro não apresenta um aviso para essas linhas.
### Tipos de lista
| Conteúdo na coluna B | Lista gerada |
|---|---|
| `ListaDebitosAposEvento` | `ListaDebitosAposEvento` |
| Qualquer outro conteúdo, inclusive vazio | `ListaDebitos` |
A identificação de `ListaDebitosAposEvento` ignora maiúsculas/minúsculas e espaços nas extremidades.
Para manter o preenchimento padronizado, utilize `ListaDebitos` nos registros normais.
### Regras dos débitos
- Cada linha reconhecida gera um item individual no JSON.
- A macro não soma valores nem elimina registros duplicados.
- `IdDebito`, `CodigoDebito` e `ValorDebito` são sempre incluídos.
- `IdEventoDebito` é incluído somente em registros após evento e quando preenchido.
- `PaDebito` não é incluído no grupo `RetPagamentoUnificado`.
- Grupos sem registros são omitidos.
- `BalancoLucroReal` é incluído no objeto `Debitos` somente quando `TributacaoLucro` contém exatamente `"1"`.
## Arquivos sem movimento
Quando `SemMovimento` é verdadeiro:
- O JSON contém `PeriodoApuracao` e `DadosIniciais`.
- O objeto `Debitos` não é incluído.
- A aba `MIT_Debitos` não é percorrida para montar os registros.
- Os campos `TributacaoLucro`, `VariacoesMonetarias` e `RegimePisCofins` são omitidos.
- Os dados do responsável continuam sendo incluídos conforme as regras de preenchimento.
## Estrutura do JSON
Para arquivos com movimento, a organização é:
```text
Objeto principal
├── PeriodoApuracao
│   ├── MesApuracao
│   └── AnoApuracao
├── DadosIniciais
│   ├── SemMovimento
│   ├── QualificacaoPj
│   ├── TributacaoLucro [condicional]
│   ├── VariacoesMonetarias [condicional]
│   ├── RegimePisCofins [condicional]
│   └── ResponsavelApuracao
│       ├── CpfResponsavel
│       ├── TelResponsavel [condicional]
│       ├── EmailResponsavel [condicional]
│       └── RegistroCrc [condicional]
└── Debitos
    ├── BalancoLucroReal [condicional]
    └── Grupos com registros
        ├── ListaDebitos [quando houver]
        └── ListaDebitosAposEvento [quando houver]
```
## Pasta e nome do arquivo
A pasta de saída está definida no código:
```text
C:\MIT\
```
Se não existir, a macro tenta criá-la com `MkDir`.
O nome segue o padrão:
```text
[CNPJ raiz]-MIT-[AAAAMM].json
```
Exemplo ilustrativo:
```text
12345678-MIT-202601.json
```
O CNPJ raiz corresponde aos oito primeiros dígitos do CNPJ informado em `MIT_Cabecalho!O2`.
A competência do nome é montada com o ano e o mês informados em `B2` e `A2`.
Se já existir um arquivo com o mesmo nome, ele será sobrescrito. Não há confirmação nem criação automática de versões.
## Gravação do arquivo
A rotina utiliza `ADODB.Stream` para salvar o conteúdo em UTF-8.
Os valores de débito são formatados com duas casas decimais e separador decimal por ponto no JSON.
A função `EscapeJson` trata:
- Barras invertidas;
- Aspas duplas;
- Quebras de linha.
Essa função é aplicada ao e-mail e ao número do registro CRC.
## Como utilizar
1. Prepare as abas `MIT_Cabecalho` e `MIT_Debitos` no arquivo que contém a macro.
2. Preencha a linha 2 da aba `MIT_Cabecalho`.
3. Para arquivos com movimento, preencha os débitos a partir da linha 2 da aba `MIT_Debitos`.
4. Confira o caminho de saída definido em `caminhoPasta`.
5. Execute a macro `GerarMIT_JSON`.
6. Localize o arquivo na pasta de saída.
7. Confira o conteúdo e valide a aceitação no sistema de destino antes de concluir o processo.
A mensagem “Arquivo gerado com sucesso” confirma a conclusão da gravação local, não a aceitação ou transmissão do arquivo.
## Funções e responsabilidades
| Rotina | Responsabilidade |
|---|---|
| `GerarMIT_JSON` | Coordena a leitura, as verificações e a gravação |
| `MontarPeriodoApuracao` | Monta o período de apuração |
| `MontarDadosIniciais` | Monta as informações iniciais e do responsável |
| `MontarDebitos` | Organiza os grupos de débitos |
| `MontarGrupoDebito` | Separa os itens entre lista normal e após evento |
| `MontarItemDebito` | Monta os campos de cada débito |
| `ToBoolean` | Interpreta os valores booleanos |
| `ApenasDigitos` | Remove caracteres que não sejam dígitos de 0 a 9 |
| `EscapeJson` | Trata caracteres de texto usados na montagem do JSON |
| `SalvarTextoUTF8` | Grava o arquivo na pasta de destino |
## Validações existentes
A macro interrompe o processamento com mensagem quando:
- O CNPJ do contribuinte não possui 14 dígitos após a limpeza.
- `VariacoesMonetarias` está vazio em um arquivo com movimento.
Essas verificações não substituem a validação completa do conteúdo.
## Limitações e cuidados
- Não há tratamento geral de erros para abas inexistentes, conversões inválidas, permissões de pasta ou falhas na gravação.
- Não há validação de dígitos verificadores de CPF ou CNPJ.
- Não há validação explícita do intervalo do mês ou da consistência do ano.
- Os campos inseridos sem aspas precisam conter representações numéricas válidas. Campos obrigatórios vazios podem produzir JSON inválido.
- A função `ApenasDigitos` remove qualquer caractere que não esteja entre `0` e `9`, incluindo letras.
- Identificadores com zeros à esquerda devem ser preservados na planilha; a macro não recompõe zeros perdidos.
- O valor da coluna N da aba `MIT_Debitos` é convertido diretamente com `CDbl`, sem verificação prévia com `IsNumeric`.
- Não há verificação de duplicidade de débitos ou de identificadores.
- Não há conferência automática de totais.
- O JSON é montado por concatenação de texto, sem validação sintática posterior.
- A macro não valida códigos, combinações de campos ou exigências condicionais do leiaute de destino.
- O caminho de saída é fixo e os arquivos existentes com o mesmo nome são sobrescritos.
## Privacidade
Não publique planilhas preenchidas ou arquivos JSON reais que contenham CPF, telefone, e-mail, dados cadastrais ou valores fiscais.
Para exemplos no repositório, utilize dados fictícios.
