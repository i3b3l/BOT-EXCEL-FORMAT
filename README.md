# BotPag - Formatador de Planilhas

Ferramenta complementar para preparar planilhas de referência utilizadas pelos robôs de processamento de pagamentos.

## Funcionalidades
- Padroniza CPF, valor, banco, agência, conta e datas.
- Preenche campos repetidos em blocos na aba Entrada.
- Mantém determinadas abas de processamento somente com cabeçalhos.
- Gera Pagamentos_Tratado.xlsx sem sobrescrever Pagamentos.xlsx.

## Como usar
1. Mantenha Pagamentos.xlsx na mesma pasta do programa.
2. Feche Pagamentos_Tratado.xlsx se estiver aberto no Excel.
3. Execute Iniciar_BotPag_Formatador.bat.
4. Confira o resultado antes de usar nos demais robôs.

## Arquivos
- Formatar_planilha.py: código-fonte.
- BotPag-Formatador.exe: executável para Windows.
- Iniciar_BotPag_Formatador.bat: inicializador.
- Instalar_Dependencias.bat: dependências para desenvolvimento.
- Pagamentos.xlsx: modelo de referência sem registros de pagamento.

## Segurança
A planilha publicada foi preparada sem registros de pagamento. Não publique planilhas operacionais com dados pessoais ou bancários reais em repositórios públicos.
