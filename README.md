Instalador de Impressoras de Rede via Script (VBS)

<img width="296" height="144" alt="image" src="https://github.com/user-attachments/assets/4a487d32-c7c3-45df-8881-2877f2d676bf" />

Este repositório contém um script VBScript (InstalaImpressoras.vbs) projetado para automatizar a instalação de impressoras de rede em computadores com sistema operacional Windows.

O script é ideal para administradores de sistemas ou equipes de TI que precisam padronizar a instalação de impressoras em múltiplos computadores de forma rápida e com feedback claro para o usuário.

Descrição
O script InstalaImpressoras.vbs executa as seguintes ações:

Lê uma lista pré-configurada de impressoras de rede.

Verifica se cada impressora da lista já está instalada no computador.

Se a impressora não estiver instalada, ele tenta a conexão.

Fornece um feedback visual (caixas de mensagem) para o usuário sobre o status de cada impressora:

Instalada com sucesso.

Já estava instalada.

Falha na instalação, exibindo o motivo do erro.

Opcionalmente, pode definir a última impressora da lista como a impressora padrão do Windows.

Exibe um relatório final com o total de novas impressoras instaladas.

Este método é muito superior a um script simples que apenas tenta adicionar as impressoras, pois evita erros desnecessários e informa o usuário e a equipe de TI sobre qualquer problema, facilitando a solução.

Funcionalidades
Instalação em Lote: Instale múltiplas impressoras de uma só vez.

Verificação de Duplicidade: Evita tentar instalar uma impressora que já existe.

Tratamento de Erros Detalhado: Em caso de falha, informa o código e a descrição do erro (ex: "O servidor de impressão não foi encontrado", "Acesso negado").

Configuração Fácil: A lista de impressoras é facilmente editável dentro do próprio arquivo.

Definir Impressora Padrão: Funcionalidade opcional para definir uma impressora como padrão do sistema.

Interativo: Mantém o usuário informado durante todo o processo através de caixas de diálogo.

Leve e Portátil: Não requer instalação de software adicional, funcionando nativamente no Windows.

Requisitos
Sistema Operacional Windows (Windows 7, 8, 10, 11, Windows Server).

Acesso de rede ao servidor de impressão (Print Server).

Permissões de usuário para instalar impressoras no computador local.

Como Usar
Siga os passos abaixo para configurar e executar o script.

1. Configuração
Antes de executar, você precisa editar o script para adicionar as impressoras da sua rede.

Abra o arquivo InstalaImpressoras.vbs com um editor de texto (como o Bloco de Notas).

Localize a seção de --- CONFIGURAÇÃO ---.

VBScript

' --- CONFIGURAÇÃO ---
' Adicione ou remova os caminhos de rede (UNC) das impressoras na lista abaixo.
Dim arrPrinters
arrPrinters = Array("\\printserver\Tecnico", "\\printserver\controladoria", "\\printserver\Marketing")

' Defina se deseja que a última impressora da lista se torne a padrão.
' Mude para True se desejar ativar esta funcionalidade.
Const SET_LAST_AS_DEFAULT = False
' --------------------
Altere a lista de impressoras: Modifique a linha arrPrinters = Array(...) com os caminhos de rede (UNC Path) das suas impressoras. Separe cada impressora com uma vírgula.

Exemplo: arrPrinters = Array("\\SRV-PRINT01\FINANCEIRO", "\\SRV-PRINT01\RH_COLOR")

(Opcional) Definir impressora padrão: Se você deseja que a última impressora da sua lista seja definida como a padrão do Windows, mude a seguinte linha:

De: Const SET_LAST_AS_DEFAULT = False

Para: Const SET_LAST_AS_DEFAULT = True

Salve e feche o arquivo.

2. Execução
Para rodar o script, simplesmente dê um duplo-clique no arquivo InstalaImpressoras.vbs.

O script exibirá mensagens informando o progresso da instalação para cada impressora e um relatório no final.

Licença
Este projeto é de código aberto e está licenciado sob a Licença MIT. Sinta-se à vontade para usar, modificar e distribuir.
