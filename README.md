# パワポ変換

## 概要

パワポのテンプレートを読み込み、出力するプロジェクト

## 環境

- windows
- python 3.7

## 環境準備

### 事前にインストールするもの

- pyhon のインストール
- virtualenv のインストール

## 環境構築と実行

```bash
cd project_dir # クローンしたプロジェクトに移動 README.mdと同階層

# 初回のみ
python -m venv venv

# 次回から
venv\Scripts\activate
pip install -r requirements.txt

deactivate # 必要に応じて、venvから抜ける

# 実行
python src/pptx-sample.template.py
```

## 実行結果

import file
- assets/input/**

export file
- assets/output/**

