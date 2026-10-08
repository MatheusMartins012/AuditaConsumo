# Controle de Avarias e Consumo

Sistema local em Flask + SQLite para registrar avarias e consumos.

## Nova função: foto da avaria

- A foto é obrigatória quando o tipo escolhido for **Avaria**.
- O campo de foto não aparece para **Consumo**.
- As fotos ficam na pasta `fotos_avarias`.
- O banco SQLite guarda o nome da foto relacionado ao lançamento.
- Na listagem aparece **📥 Baixar foto** para as avarias que possuem foto.
- Na edição é possível substituir uma foto existente.
- Se uma Avaria for alterada para Consumo, a foto antiga é removida.
- Formatos aceitos: JPG, JPEG, PNG e WEBP.
- Tamanho máximo: 10 MB.

## Instalação no Windows

Abra o PowerShell dentro da pasta do projeto:

```powershell
py -m venv .venv
.venv\Scripts\activate
py -m pip install -r requirements.txt
py app.py
```

No computador:

`http://127.0.0.1:5000`

Para acessar pelo celular na mesma rede Wi-Fi:

1. No PC, abra o PowerShell e execute `ipconfig`.
2. Procure o **Endereço IPv4**.
3. No celular, estando no mesmo Wi-Fi, abra:
   `http://IP-DO-PC:5000`

Exemplo:

`http://192.168.1.10:5000`

Se o Windows perguntar sobre o Firewall, permita o acesso em rede privada.
