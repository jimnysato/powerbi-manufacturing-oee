# Ficha em PDF dentro do Power BI — pré-requisitos

Objetivo: na página "Ficha diária", o usuário escolhe data e injetora nos filtros e baixa a ficha do Report Builder em PDF, já filtrada. Hoje há só um botão para um PDF de exemplo fixo (30/09/2026), hospedado no site.

## O que é preciso

1. **Conta corporativa da Microsoft.** O Power BI Service não aceita e-mail pessoal (@icloud, @gmail). Caminho: domínio próprio (ex.: vitorsato.com.br) + Microsoft 365 Business Basic.
2. **Licença Power BI Pro** (ou Premium por Usuário) para quem publica. Em workspace comum, **quem visualiza também precisa de Pro**; só em capacidade Fabric/Premium o leitor pode ser gratuito.
3. **Workspace no Power BI Service** para publicar o modelo, o dashboard e o relatório paginado.
4. **Publicar o modelo da demo B.** Para os dados se atualizarem sozinhos, mover os Excel para OneDrive/SharePoint (sem gateway) ou instalar um gateway. Para o portfólio, com dados fixos, publicar sem atualização agendada já basta.
5. **Trocar a fonte do relatório paginado:** em vez do Power BI Desktop local (`OLEDB-MD`, porta que muda), apontar para o modelo publicado (`PBIDATASET`), e publicar o `.rdl` no mesmo workspace.
6. **Adicionar o parâmetro Injetora** ao relatório paginado (hoje só tem Data), para receber a injetora escolhida no filtro.
7. **Na página Ficha diária:** inserir o visual **"Relatório paginado"**, escolher a ficha, ligar os parâmetros Data e Injetora aos campos dos filtros e ativar a barra de ferramentas (ela tem o botão de exportar PDF).
8. **Configuração do administrador:** exportar para PDF e o visual de relatório paginado precisam estar permitidos nas configurações do locatário (vêm permitidos por padrão).

## Limitações

- **"Publicar na Web" (link público, sem login) não mostra o visual de relatório paginado.** Para o portfólio público, continuam valendo os prints, o vídeo e o PDF de exemplo.
- Os custos (domínio, Microsoft 365, licença Pro) mudam com frequência: conferir os valores na hora de contratar.
