# Painel VM Shatan

Site de aprovação das fotos das mesas OPPO e painel das lojas.

- `public/` — o site (o Cloudflare publica esta pasta sozinho a cada mudança).
- `apps-script/` — cópia dos scripts da planilha "Controle de Mesas" (Sincronizar, Feishu, Painel). Não têm senhas: as chaves ficam nas Propriedades do script.
- `wrangler.jsonc` — configuração do Cloudflare.

Senhas e chaves (App Secret do Feishu, service_role do Supabase) **nunca** entram aqui.
