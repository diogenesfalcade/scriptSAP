# Rotina-Python

### Bot de extração e limpeza de dados do SAP, com integração e uso do VBScript
### O processo é acessar o SAP com uma automação que acesse o sistema RP localmente
### Dentro do repositório encontramos tanto o bot (baseado em os), quanto os scripts de extração (VBScript), além da limpeza (Pandas & Python Puro)

### Você deve apenas rodar o script através da main e a estrutura de pastas deve ser mantida igual
### O pré-requisito para executar o script é manter a pasta do Teams de PMO & Planning Team sincronizada ao seu pc,
### para que seja possível acessar diretamente pelo Windows Explorer

### NOVOS PROJETOS
### 1. Adicionar o Project ID (6 digitos - NNNNNN) seguido de um 'tab'(\t) com o nome do projeto - caso contrário o bot não irá identificar
### 2. Classificar TODOS os pacotes de trabalho com o WBS (NNNNNN-XXBR) seguido de 'tab'(\t) com seu ProjectID (6 digitos - NNNNNN) e por fim, 
### 	também seguido de tab, o nome da célula, que deve seguir exatamente o nome inserido nas outras células dos outros projetos