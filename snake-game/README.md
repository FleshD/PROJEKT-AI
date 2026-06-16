# Had v prohlížeči

Jednoduchá hra Had napsaná pouze v HTML, CSS a JavaScriptu.

## Spuštění

Otevři `snake-game/index.html` přímo v prohlížeči, nebo spusť lokální HTTP server z kořene repozitáře:

```bash
python3 -m http.server 8000
```

Potom otevři <http://127.0.0.1:8000/snake-game/>.

## Ověření Playwrightem

Nainstaluj závislosti v adresáři `snake-game` a spusť smoke test, který nastartuje lokální server, otevře hru a uloží screenshot do `snake-game/screenshot.png`:

```bash
cd snake-game
npm install
npm run verify
```

## Ovládání

- Šipky nebo klávesy `W`, `A`, `S`, `D` mění směr hada.
- Tlačítko **Start / restart** spouští novou hru.
- Na mobilu fungují dotyková tlačítka pod herní plochou.
- Nejlepší skóre se ukládá do `localStorage` prohlížeče.
