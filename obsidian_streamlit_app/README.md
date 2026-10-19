# 黒曜石EDXRF産地判別・クラスター解析プログラム

## GitHubへの配置とStreamlitへの登録

1. このフォルダの `app.py`、`analysis_core.py`、`requirements.txt` を、GitHubリポジトリの同じフォルダにアップロードします。
2. Streamlit Community Cloudで「Create app」を選択します。
3. Repositoryにアップロード先のリポジトリ、Branchにそのブランチを指定します。
4. Main file pathは、リポジトリ直下に置いた場合は `app.py`、`obsidian_streamlit` フォルダに置いた場合は `obsidian_streamlit/app.py` を指定します。
5. Python 3.12を選択してDeployを実行します。

既存アプリのapp.pyと重ならないよう、新しいリポジトリ、または専用フォルダに配置してください。研究用ExcelをGitHubにアップロードする必要はありません。起動したアプリの画面で選択します。

## 使用方法

Excelで個別判定を実行し、再計算・保存したxlsxをアップロードして「解析を実行」を押します。Medianシートの代表値と、AS-1以降のシートのL4:T7にあるD1〜D3の楕円パラメータを使用します。

Rb*、Sr*、Zr*は100×元素値÷(Rb+Sr+Y+Zr)です。標準化後にHDBSCAN（min_cluster_size=3、min_samples=3）でクラスターを検出します。C1から資料数の多い順に番号を付けます。クラスター楕円は95%分布楕円です。

結果Excelは従来と同じ5シートです。Sample_Cluster_Resultsは上下左右中央揃えです。散布図は画面上で表示し、Excelには追加しません。画面で重ねる化学組成グループを変更しても、比較計算と出力Excelは全グループを対象にします。

内包率は資料数の割合であり、楕円同士の面積の重なり率や所属確率ではありません。資料別に表示する候補もクラスター全体の候補です。入力したMedianの選別方法の影響を受けます。

データはセッション内で扱い、一時ファイルは処理終了時に削除します。公開サーバーでは解析のためにExcelがそのサーバーへ送信されます。

## ローカル実行

```sh
python -m pip install -r requirements.txt
streamlit run app.py
```

Streamlit公式資料: https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app
