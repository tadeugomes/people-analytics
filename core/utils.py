#!/usr/bin/env python3
"""
Core utilities shared across pipelines: normalization, mappings, indices,
thresholds/config loading, and basic HOME templating helpers.
"""
from __future__ import annotations

import json
import os
import unicodedata
from typing import Dict, Tuple, Any

import numpy as np
import pandas as pd


def normalize(s: Any) -> str:
    try:
        s2 = unicodedata.normalize('NFKD', str(s)).encode('ASCII', 'ignore').decode('ASCII')
        return ''.join(ch if ch.isalnum() else '_' for ch in s2.lower()).strip('_')
    except Exception:
        return str(s).lower()


def load_config() -> Dict[str, Any]:
    """Load diversity config from env or common paths.

    Priority:
      - env DIVERSITY_CONFIG
      - ./config_diversidade.json
      - ./docs/config_diversidade.json
    """
    candidates = [
        os.environ.get('DIVERSITY_CONFIG'),
        os.path.join(os.getcwd(), 'config_diversidade.json'),
        os.path.join(os.getcwd(), 'docs', 'config_diversidade.json'),
        os.path.join(os.getcwd(), 'docs', 'config_example.json'),
    ]
    for path in candidates:
        if path and os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                continue
    return {}


def get_thresholds(cfg: Dict[str, Any] | None = None) -> Tuple[float, float]:
    cfg = cfg or load_config()
    thr = cfg.get('thresholds', {}) if isinstance(cfg, dict) else {}
    low = float(thr.get('low', 0.6))
    high = float(thr.get('high', 0.8))
    return low, high


def standardize_gender(series: pd.Series) -> pd.Series:
    mapping = {
        'm': 'Masculino', 'masc': 'Masculino', 'masculino': 'Masculino', 'homem': 'Masculino', 'male': 'Masculino', 'man': 'Masculino',
        'f': 'Feminino', 'fem': 'Feminino', 'feminino': 'Feminino', 'mulher': 'Feminino', 'female': 'Feminino', 'woman': 'Feminino'
    }
    def one(x):
        if pd.isna(x):
            return 'Outro/NS'
        nx = normalize(x)
        if nx in mapping:
            return mapping[nx]
        return 'Masculino' if nx in ['h'] else ('Feminino' if nx in ['w'] else 'Outro/NS')
    return series.apply(one)


def standardize_race(series: pd.Series) -> pd.Series:
    mapping = {
        'branca': 'Branca', 'branco': 'Branca',
        'preta': 'Preta', 'preto': 'Preta', 'negra': 'Preta', 'negro': 'Preta',
        'parda': 'Parda', 'amarela': 'Amarela',
        'indigena': 'Indígena', 'indigena': 'Indígena', 'indígena': 'Indígena',
        'nao_informado': 'Não informado', 'nao_declarado': 'Não informado', 'nd': 'Não informado', 'ns': 'Não informado', 'nr': 'Não informado'
    }
    def one(x):
        if pd.isna(x):
            return 'Não informado'
        nx = normalize(x)
        return mapping.get(nx, 'Não informado')
    return series.apply(one)


def simpson_index(counts: pd.Series | Dict[Any, int]) -> float:
    if isinstance(counts, pd.Series):
        total = counts.sum()
        if total == 0:
            return 0.0
        p2 = ((counts / total) ** 2).sum()
        return float(1 - p2)
    else:
        total = sum(counts.values())
        if total == 0:
            return 0.0
        return float(1 - sum((c/total)**2 for c in counts.values()))


def shannon_index(counts: pd.Series | Dict[Any, int]) -> float:
    if isinstance(counts, dict):
        counts = pd.Series(counts)
    total = counts.sum()
    if total == 0:
        return 0.0
    p = counts / total
    p = p[p > 0]
    return float(-(p * np.log(p)).sum())


def interpret_index(idx: float, scope: str = 'geral') -> str:
    low, high = get_thresholds()
    if idx >= high:
        return f'Alta diversidade de {scope} (índice = {idx:.3f}).'
    if idx >= low:
        return f'Diversidade moderada de {scope} (índice = {idx:.3f}).'
    return f'Baixa diversidade de {scope} (índice = {idx:.3f}).'


def find_gender_column(df: pd.DataFrame) -> str | None:
    for c in df.columns:
        nc = normalize(c)
        if any(k in nc for k in ['genero', 'gnero', 'sexo', 'gender']):
            return c
    return None


def find_race_column(df: pd.DataFrame) -> str | None:
    for c in df.columns:
        nc = normalize(c)
        if any(k in nc for k in ['raca', 'raça', 'cor', 'race', 'etnia', 'ethnic']):
            return c
    return None

