#!/usr/bin/env python3
"""Confere se o script da página sobrevive ao minificador do Apps Script.

O HtmlService remove comentários ao servir a página e trata qualquer "/*" ou "//"
como comentário, mesmo dentro de strings e atributos (ex.: accept="image/*",
'https://...'). Isso corta o código e o quadro fica "carregando" para sempre.
Uso: python3 tools/checar_minificador.py src/Index.html
"""
import re, sys
s = open(sys.argv[1] if len(sys.argv) > 1 else 'src/Index.html', encoding='utf-8').read()
problemas = []
for m in re.finditer(r'<script>([\s\S]*?)</script>', s):
    base = s[:m.start()].count('\n') + 1
    em_bloco = False
    for k, linha in enumerate(m.group(1).split('\n')):
        t = linha.strip()
        if em_bloco:
            if '*/' in t: em_bloco = False
            continue
        if t.startswith('/*'):
            if '*/' not in t[2:]: em_bloco = True
            continue
        if t.startswith('//') or t.startswith('*'): continue
        codigo = re.sub(r'\s//\s.*$', '', linha)  # comentário de fim de linha " // ..."
        if '/*' in codigo or re.search(r'(?<!\\)//', codigo.replace(':\\/\\/', '')):
            problemas.append(f'{base + k}: {t[:120]}')
if problemas:
    print('❌ Trechos que o minificador do Apps Script vai cortar (use \\/\\/ ou evite "/*"):')
    print('\n'.join(problemas)); sys.exit(1)
print('✅ Script seguro para o minificador do Apps Script')
