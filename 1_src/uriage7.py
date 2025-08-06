import pandas as pd

# ユーザーからの入力 (注文日)
order_date = input("注文日をyyyymmddで入力してください (例: 20250626): ")
if not order_date.isdigit() or len(order_date) != 8:
    print("エラー: 日付はyyyymmdd形式の数字8桁で入力してください。")
    exit()

# --- ファイルパス定義 ---
FILE_DAILY_ORDER = "uriage020_hibi.csv"  # 日々の注文
FILE_ITEM_MASTER = "uriage010_nemu.csv"  # お弁当マスタ
FILE_MONTHLY_ACCUM = "uriage030_ruiseki.csv" # 月次累計

# --- 設定 (必要に応じて) ---
pd.set_option('display.unicode.east_asian_width', True) # 日本語表示の調整

# --- 1. データ読み込みと前処理 ---
# encoding は "cp932" で固定します。
try:
    df_daily_order = pd.read_csv(FILE_DAILY_ORDER, encoding="cp932")
    df_item_master = pd.read_csv(FILE_ITEM_MASTER, encoding="cp932")
    df_monthly_accum = pd.read_csv(FILE_MONTHLY_ACCUM, encoding="cp932")
except FileNotFoundError as e:
    print(f"エラー: 必要なファイルが見つかりません。パスを確認してください: {e}")
    exit() # ファイルがない場合は処理を中断

# 型変換 (お品コードとコードを結合キーにするためInt64に統一)
df_daily_order['お品コード'] = df_daily_order['お品コード'].astype('Int64').fillna(0) # 欠損値は0に
df_item_master['コード'] = df_item_master['コード'].astype('Int64').fillna(0) # 欠損値は0に
# 月次累計ファイルも金額を念のため整数型に。
if '金額' in df_monthly_accum.columns:
    df_monthly_accum['金額'] = pd.to_numeric(df_monthly_accum['金額'], errors='coerce').fillna(0).astype('Int64')




# 日々の注文とお弁当マスタの結合
df_todays_orders = pd.merge(
    df_daily_order,
    df_item_master,
    how="inner",
    left_on=["お品コード"],
    right_on=["コード"]
)

# 金額の計算と**整数型への変換、欠損値の0埋め**
df_todays_orders["年月日"] = order_date # 入力された注文日をセット
df_todays_orders["金額"] = df_todays_orders['数量'] * df_todays_orders['値段']
# ここで金額を整数型に変換し、NaNは0にします。
df_todays_orders['金額'] = df_todays_orders['金額'].fillna(0).astype('Int64')

# 必要最低限の列に絞る（お弁当注文と集計に必要な列のみ）
print("\n--o--i--k--a--w--a-- 当日分オーダーリスト --m--a--s--a--h--i--k--o--")

su = df_todays_orders["金額"].sum()

# 表示前にフォーマットを設定し、表示後にリセット
pd.options.display.float_format = '{:.0f}'.format
df_todays_orders_display = df_todays_orders[['年月日', '名前', '部門', 'コード', '名称', '数量', '値段', '金額', '集計キー', 'お店']]

df_todays_orders_display = df_todays_orders_display.sort_values(by=['お店', '名前'], ascending=True)

#######################################################

df_todays_orders = df_todays_orders_display.reset_index(drop=True)

#######################################################

print(df_todays_orders.fillna(0))#
pd.reset_option('display.float_format') # 表示設定をリセット
df_todays_orders.to_csv('df_todays_orders_display.csv', encoding="cp932", index=True)#

# 合計金額の表示
print(f"合計は、{su:.0f}円です。")


print("\n--- 部門別個人別集計 ---")
# 個人別集計 (金額)
# margins=True を削除し、手動で合計行・列を追加
df_personal_summary = df_todays_orders.pivot_table(
    index='名前',
    columns='部門',
    values='金額',
    aggfunc='sum'
)

# 欠損値を0で埋める（合計計算前に実施）
df_personal_summary = df_personal_summary.fillna(0)

# 行の合計（名前ごとの合計）を追加
df_personal_summary['合計'] = df_personal_summary.sum(axis=1)

# 列の合計（部門ごとの合計）を追加
# まず合計行を計算
col_sums = df_personal_summary.sum(axis=0).to_frame().T
col_sums.index = ['合計']
df_personal_summary = pd.concat([df_personal_summary, col_sums])

# 表示前にフォーマットを設定し、表示後にリセット
pd.options.display.float_format = '{:.0f}'.format
print(df_personal_summary) # fillna(0)は先に実施済みなので不要
pd.reset_option('display.float_format') # 表示設定をリセット

# CSV保存時は数値のまま保存される
df_personal_summary.to_csv('Department_Personal_Summary.csv', encoding="cp932")
print("「Department_Personal_Summary.csv」を作成しました。")

# 会社負担分の考慮:
# ここは、会社負担額の「ルール」がコードに組み込まれていないため、
# 現時点では集計のみです。


print("\n--- お弁当注文 FAX 様式 ---")

# '集計キー'のユニークな値を取得
unique_集計キー = df_todays_orders['集計キー'].unique()

for key in unique_集計キー:
    print(f"\n--- 集計キー: {key} のFAX様式を作成中 ---")

    # 現在の集計キーに該当するデータのみをフィルタリング
    df_filtered_orders = df_todays_orders[df_todays_orders['集計キー'] == key]

    # お弁当名と数量の集計
    df_fax_summary_per_key = df_filtered_orders.pivot_table(
        index=['名称', '値段'],
        columns='お店',
        values='数量',
        aggfunc='sum',
        margins=False
    )

    # 表示前にフォーマットを設定し、表示後にリセット
    pd.options.display.float_format = '{:.0f}'.format
    print(df_fax_summary_per_key.fillna(0))
    pd.reset_option('display.float_format') # 表示設定をリセット

    # FAX送信用CSVとして保存
    output_filename = f'FAX_Order_Summary_{key}.csv'
    df_fax_summary_per_key.to_csv(output_filename, encoding="cp932")
    print(f"「{output_filename}」を作成しました。")

print("\n--- 全てのFAX様式ファイルの作成が完了しました。 ---")


print("\n--- 月次累計データへの追加 ---")
# 今日の注文データを月次累計ファイルに追記
# df_todays_orders の列順を df_monthly_accum に合わせる
df_todays_orders_for_accum = df_todays_orders[['年月日', '名前', '部門', 'コード', '名称', '数量', '値段', '金額', 'お店', '集計キー']].copy()


import os

# 実行したいVBScriptファイルのパス
vbs_file = "010_hiraku3.vbe"

# ファイルが存在するか確認する（オプション）
if not os.path.exists(vbs_file):
    print(f"エラー: ファイル '{vbs_file}' が見つかりません。パスを確認してください。")
else:
    try:
        os.startfile(vbs_file)#ダブルクリックしたのと同じような効果
    except Exception as e:
        print(f"'{vbs_file}' の実行中にエラーが発生しました: {e}")


S = input("OKならEnterを押してください。NGなら右上×を押してください。")#ワンクッション置く

# 結合して新しいDataFrameを作成し、CSVに保存
df_updated_monthly_accum = pd.concat([df_monthly_accum, df_todays_orders_for_accum], ignore_index=True)
df_updated_monthly_accum.to_csv(FILE_MONTHLY_ACCUM, encoding="cp932", index=False)
print(f"「{FILE_MONTHLY_ACCUM}」に今日の注文データを追加しました。")

# 月次合計の表示例 (デバッグ用や確認用)
monthly_total_amount = df_updated_monthly_accum[df_updated_monthly_accum['年月日'].str[:6] == order_date[:6]]['金額'].sum()
# 合計金額の表示
print(f"当月累計金額 (現時点): {monthly_total_amount:.0f}円")


print("\n--- 日々注文ファイルの初期化 ---")

# df_daily_order の「お品コード」列を全て空文字列に設定
df_daily_order['お品コード'] = ""

# df_daily_order の「数量」を全て1に設定
df_daily_order['数量'] = 1

# 変更したDataFrameをuriage020_hibi.csvに上書き保存
df_daily_order.to_csv(FILE_DAILY_ORDER, encoding="cp932", index=False)

print(f"「{FILE_DAILY_ORDER}」を初期化しました。")

print("\n--- 全ての処理が完了しました。 ---")
