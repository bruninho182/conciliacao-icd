# Conciliação ICD

![Ícone da extensão](icon-128.png)

Extensão para Chrome que consulta reservas no ICD e preenche vouchers em uma cópia de uma planilha Excel.

## Instalação

1. Baixe e descompacte este repositório ou clone-o com Git.
2. Abra `chrome://extensions` e ative o **Modo do desenvolvedor**.
3. Clique em **Carregar sem compactação** e selecione a pasta que contém o arquivo `manifest.json`.

## Como usar

1. Entre no ICD e abra **Vendas → Reservas**. Deixe o campo **Pesquisar** visível e confira os filtros.
2. Clique no ícone da extensão e selecione o arquivo `.xlsx`.
3. Clique em **Testar 5 consultas**, baixe a cópia e confira os vouchers.
4. No mesmo painel, clique em **Consultar restantes**. Ao terminar, use **Baixar cópia** e, se houver linhas pendentes, **Baixar pendências**.

A planilha deve ter a aba **Relatório Imply**, com os cabeçalhos **Status**, **Voucher** e **Localizador** nas colunas A, B e C da linha 4. A extensão pula os registros com status **Cancelado (CA)** e as linhas com voucher já preenchido. Ela pesquisa o trecho final do UUID e só grava o código quando a tela confirma um único resultado com o localizador completo.

Os arquivos são processados no Chrome. A extensão usa a tela de Reservas com a sessão do usuário e pausa se identificar uma verificação ou restrição. Baixe uma cópia antes de fechar ou atualizar a aba para salvar o progresso.
