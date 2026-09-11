# OftalmoCastro — Gestão de Estoque de Lentes

Sistema React/Vite para controle de lentes intraoculares.

## Incluído
- Dashboard profissional
- Cadastro, edição e exclusão de lentes
- Nome, fabricante, modelo, potência/numeração, tipo, lote, número de série, validade, quantidade, estoque mínimo, localização e fornecedor
- Entrada, saída, reserva, devolução e ajuste de inventário
- Data, horário, responsável, motivo, paciente e cirurgia em cada movimentação
- Reserva de lente por paciente, cirurgia, médico, data, horário e olho
- Alertas de estoque mínimo e validade
- Fornecedores
- Usuários e status de acesso
- Relatórios e exportação CSV
- Histórico completo
- Persistência local no navegador
- Layout responsivo

## Como rodar
1. Instale Node.js 20+
2. Abra a pasta no terminal
3. Rode `npm install`
4. Rode `npm run dev`

## Para uso definitivo na clínica
Esta versão salva os dados no navegador. Para uso real em vários computadores ao mesmo tempo, o próximo passo é conectar a um backend com banco PostgreSQL/Supabase, login, permissões, backup, auditoria e sincronização em nuvem.
