#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Métricas da fonte usada no documento.

Serve para garantir que nenhuma coluna de tabela seja estreita demais para
a maior palavra indivisível que nela aparece — situação que faria o texto
ser quebrado no meio da palavra, sem hífen.
"""

from fontTools.ttLib import TTFont

ARQUIVOS = {
    False: "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    True: "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
}

# Recuo interno das células, somados os dois lados, em centímetros.
MARGEM_CELULA_CM = 0.30
MARGEM_CELULA_TWIPS = 85


def _carregar(caminho):
    fonte = TTFont(caminho)
    return fonte["head"].unitsPerEm, fonte["hmtx"], fonte.getBestCmap()


_CACHE = {}


def _metricas(negrito):
    if negrito not in _CACHE:
        _CACHE[negrito] = _carregar(ARQUIVOS[negrito])
    return _CACHE[negrito]


def largura_cm(texto, tamanho_pt, negrito=False):
    """Largura do texto em centímetros, sem quebra de linha."""
    upm, hmtx, cmap = _metricas(negrito)
    substituto = cmap.get(ord("o"))
    total = 0
    for ch in texto:
        nome = cmap.get(ord(ch), substituto)
        total += hmtx[nome][0]
    return total / upm * tamanho_pt / 72.0 * 2.54


def maior_palavra(texto):
    """Maior sequência indivisível do texto, considerando o hífen e a
    barra como pontos legítimos de quebra."""
    maior = ""
    for bruto in texto.replace("\n", " ").split(" "):
        pedaco = bruto
        for separador in ("-", "/", "—", "·"):
            pedaco = pedaco.replace(separador, separador + " ")
        for parte in pedaco.split(" "):
            if len(parte) > len(maior):
                maior = parte
    return maior


def largura_minima_cm(texto, tamanho_pt, negrito=False):
    """Largura mínima da coluna para que o texto não seja partido."""
    return largura_cm(maior_palavra(texto), tamanho_pt, negrito)
