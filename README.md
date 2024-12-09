# Projeto de Extração e Processamento de Dados de Presos (Canaime)

Este projeto tem como objetivo extrair, processar e armazenar dados de presos cadastrados em um determinado sistema (Canaime). A partir de uma página de pesquisa online, são obtidos os IDs e nomes dos reeducandos, filtrados conforme critérios específicos, e finalmente, os resultados são salvos em um arquivo Excel.

## Pré-requisitos

- Python 3.7 ou superior  
- Acesso à Internet (para acessar as páginas de pesquisa)  
- Ambiente virtual (opcional, mas recomendado)  
- Módulos necessários listados em `requirements.txt` (por exemplo: `openpyxl`, `login_canaime`)

## Instalação

1. **Clonar o repositório**:
   ```bash
   git clone https://github.com/A-Assuncao/pesquisa-saida-PAMC
   cd pesquisa-saida-PAMC
2.  **Criar um ambiente virtual (opcional, mas recomendado)**:

    ```bash
    python -m venv venv
     ```
    -   Linux/MacOS:
        ```bash
        source venv/bin/activate
        ```
    -   Windows:
  
        ```bash
        venv\Scripts\activate
        ```
        
3.  **Instalar as dependências**:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Instalar Chromium**:
    ```bash
    playwright install chromium
    ```

## Como Executar

```bash
python main.py
```

Caso o projeto esteja organizado de outra forma, ajuste o comando conforme necessário.

## Estrutura do Projeto

Exemplo de organização:
```plaintext
pesquisa-saida-PAMC/
├── login_canaime.py        # Módulo responsável pelo login e navegação
├── main.py                 # Script principal (com as funções atualizadas)
├── lista_ids_saida.json    # Gerado durante a execução
├── presos_saida.xlsx       # Gerado durante a execução
└── requirements.txt        # Lista de dependências
```

**Observação**: O arquivo `login_canaime.py` não é fornecido. Ele deve conter a classe `Login` para autenticação e navegação no site alvo.

## Funcionamento

1.  **Geração da Lista de IDs**:  
    A função `lista_ids_saida` acessa a página de pesquisa, extrai a quantidade total de presos, itera pelas páginas e coleta IDs e nomes dos reeducandos que atendem ao critério "SAIDA". Os resultados são salvos em `lista_ids_saida.json`.
    
2.  **Verificação da Lista de IDs**:  
    A função `busca_dados` verifica se `lista_ids_saida.json` existe. Caso não exista, chama `lista_ids_saida` para gerar o arquivo. Caso exista, carrega os dados do arquivo diretamente.
    
3.  **Filtragem por Unidades e Datas**:  
    A função `busca_datas` acessa a página de certidão de cada preso, obtendo a última unidade e data de registro. Em seguida, filtra aqueles com unidade "PAMC" no ano de 2024, formando a lista final.
    
4.  **Exportação para Excel**:  
    A função `salvar_excel` recebe a lista filtrada e cria o arquivo `presos_saida.xlsx`, formatado com cabeçalhos e colunas ajustadas.
    

## Tratamento de Erros

-   **Acesso às páginas**: Em caso de erro ao acessar as páginas, exceções são capturadas e mensagens de erro são exibidas, evitando a interrupção total do programa.
-   **Leitura/Escrita de Arquivos**: Erros ao ler ou escrever arquivos JSON e Excel são tratados, informando o usuário do problema.
-   **Dados Inesperados**: Se a estrutura da página for alterada ou dados inconsistentes forem retornados, o erro é tratado e relatado, auxiliando na correção.

## Contribuindo

1.  Faça um fork do repositório.
2.  Crie uma nova branch:
    
	```bash
	git checkout -b feature-minha-contribuicao
	```
    
3.  Faça as modificações necessárias e confirme:   
  
	```bash
	git commit -m "Implementa nova funcionalidade"
	```
    
4.  Envie as alterações:      

	```bash
	git push origin feature-minha-contribuicao
	```
    
5.  Crie um Pull Request para o repositório original.

## Licença

Este projeto está sob a licença MIT. Consulte o arquivo `LICENSE` (caso exista) para mais detalhes.
