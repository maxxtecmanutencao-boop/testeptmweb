# 🚀 PTM JSL - Sistema de Consultas 2.0

## 📋 Descrição
Sistema avançado de consultas e monitoramento de pedidos logísticos desenvolvido para a Petrobras.

## 🎯 Funcionalidades
- ✅ Dashboard Executivo com métricas em tempo real
- ✅ Busca e consulta por remessa
- ✅ Planilha completa com edição de dados
- ✅ Analytics avançado com gráficos interativos
- ✅ Análise detalhada de atrasos por status
- ✅ Rastreamento online de pedidos
- ✅ Atualização automática via SharePoint
- ✅ Sistema de backup automático

## 💻 Requisitos do Sistema
- Windows 7/8/10/11
- Python 3.8 ou superior
- Google Chrome (para rastreamento online)
- Conexão com internet

## 📦 Instalação

### Opção 1: Instalação Automática (RECOMENDADO)
1. Extraia todos os arquivos para uma pasta
2. Execute o arquivo `INSTALADOR.bat`
3. Aguarde a instalação das dependências
4. Execute `INICIAR_SISTEMA.bat` para iniciar

### Opção 2: Instalação Manual
1. Instale o Python: https://www.python.org/downloads/
2. Abra o Terminal/CMD na pasta do sistema
3. Execute: `pip install -r requirements.txt`
4. Execute: `streamlit run app_melhorado.py`

## 🚀 Como Usar

### Iniciando o Sistema
- Execute o arquivo `INICIAR_SISTEMA.bat`
- O navegador abrirá automaticamente em: http://localhost:8501

### Navegação
Use o menu lateral para navegar entre as telas:
- 🏠 **Dashboard**: Visão geral executiva
- 📊 **Resumo por Status**: Busca e consulta de remessas
- 📋 **Planilha Completa**: Visualização e edição de dados
- 📈 **Analytics**: Análises avançadas e atrasos
- 🌐 **Rastreamento Online**: Consulta em tempo real
- 🔄 **Atualizar Sistema BD**: Atualização da base de dados

## 📁 Arquivos Necessários
- `app_melhorado.py` - Código principal
- `BD.xlsx` - Base de dados
- `Petrobras.png` - Logo Petrobras
- `logo jsl.png` - Logo JSL
- `requirements.txt` - Dependências
- `INSTALADOR.bat` - Instalador automático
- `INICIAR_SISTEMA.bat` - Executável

## 🔧 Solução de Problemas

### O sistema não inicia
1. Verifique se o Python está instalado: `python --version`
2. Execute novamente o `INSTALADOR.bat`
3. Verifique se a porta 8501 está livre

### Erro ao rastrear pedidos
1. Verifique se o Google Chrome está instalado
2. Verifique sua conexão com internet
3. Tente novamente após alguns segundos

### Erro ao atualizar do SharePoint
1. Faça login no SharePoint manualmente
2. Use a Opção 2: Upload Manual
3. Baixe o arquivo BD.xlsx e faça upload

## 👨‍💻 Desenvolvedor
- **Nome**: Djalma A Barbosa (FYF9)
- **Empresa**: Petrobras
- **Gerência**: PCAD/OPARM/ARM-II
- **Versão**: 2.0
- **Data**: Janeiro/2026

## 📞 Suporte
Para dúvidas ou problemas, entre em contato com o desenvolvedor.

## 📄 Licença
Uso exclusivo Petrobras - Todos os direitos reservados.