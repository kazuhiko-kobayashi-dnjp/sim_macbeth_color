# Usage:
#   make build   -- build_interactive.py を実行してHTMLを再生成
#   make encrypt -- StatiCryptでパスワード保護HTMLを生成 (encrypted/macbeth_interactive.html)
#   make deploy  -- encrypt + git commit + push (GitHub Pages へ反映)
#
# 初回のみパスワードを設定:
#   export STATICRYPT_PASSWORD="your-secret-passphrase"

EXCEL  := 201006_車両WSカメラ影響定量化まとめ２.xlsx
SRC    := macbeth_interactive.html
OUTDIR := docs

build:
	python3 build_interactive.py --excel "$(EXCEL)" --out $(SRC)

encrypt: build
	mkdir -p $(OUTDIR)
	npx staticrypt $(SRC) \
	  --directory $(OUTDIR) \
	  --remember 7 \
	  --short \
	  --title "MacBeth Simulation"
	@echo "→ $(OUTDIR)/$(SRC)"

deploy: encrypt
	git add $(OUTDIR)/$(SRC)
	git add build_interactive.py sim_macbeth.py .gitignore Makefile
	git commit -m "update simulation"
	git push
