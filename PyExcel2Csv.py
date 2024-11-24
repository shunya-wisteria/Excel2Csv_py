# ----------------------------------------------------------
#  Excel-CSV変換
# 
# 概要：
#   xlsxファイルパスをinputに、対象ファイルを展開し、
#   シートごとに出力フォルダパスで指定された場所にcsv出力する
# 
# 引数：
# ・入力xlsxファイルパス
# ・出力フォルダパス
#  
# リターンコード：
# ・81：引数不足
# ・82：入力ファイルパス不正
# ・83：出力フォルダ作成失敗
# ・99：その他例外
# ----------------------------------------------------------

import sys
import os
import pandas as pd
import csv

# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
# % 関数定義 %
# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

# --------------------------------------
# 引数チェック
# 
# 引数：実行時引数
# 戻り値：リターンコード
#  81：引数不足、82：入力ファイルパス不正
# --------------------------------------
def CheckArgs(args):
  # 引数チェック
  if(len(args) < 3):
    print("引数が不足しています、第1引数：入力xlsxファイル、第2引数：出力フォルダパス")
    return 81
  
  # 入力ファイル存在チェック
  if not os.path.isfile(args[1]):
    print ("入力xlsxファイルが存在しません。")
    return 82

  return 0

#  --------------------------------------
#  事前処理
#  
#  引数：出力フォルダパス
#  戻り値：リターンコード
#   83：出力フォルダ作成失敗
#  --------------------------------------
def PreProc(outPath):
  try:
    # 出力フォルダが存在しない場合、作成
    if not os.path.isdir(outPath):
      os.mkdir(outPath)
  except:
    return 83
  
  return 0

#  --------------------------------------
#  エクセル-CSV変換
#  
#  引数：
#  ・入力xlsxファイルパス
#  ・出力フォルダパス
#  戻り値：出力ファイル一覧
#  --------------------------------------
def Excel2Csv(inFile, outPath):
  outFiles = []
  try:
    # エクセルファイル名
    inFileNm = os.path.basename(inFile)
    prefix = os.path.splitext(inFileNm)[0]

    # エクセルブック読込
    excelBook = pd.ExcelFile(inFile)

    # シートを走査
    for sheetName in excelBook.sheet_names:
      # シート読込
      sheet = pd.read_excel(inFile, sheet_name=sheetName)
      # 出力ファイル名作成
      outFile = outPath + "/" + prefix + "_" + sheetName + ".csv"
      # CSVファイル出力
      sheet.to_csv(outFile, index=False, quoting=csv.QUOTE_ALL)

      # 出力ファイル名を格納
      outFiles.append(outFile)

  except:
    print("Excel-CSV変換に失敗しました。")
    return None
  
  return outFiles

#  --------------------------------------
#  メイン関数
#  
#  引数：実行時引数
#  戻り値：リターンコード
#  --------------------------------------
def main(args):
  check = 0
  
  # 引数チェック
  check = CheckArgs(args)
  if(check != 0):
    return check
  
  # 入力ファイルパス
  IN_FILE = args[1]
  # 出力ファイルパス
  OUT_PATH = args[2]

  # 事前処理
  check = PreProc(OUT_PATH)
  if(check != 0):
    return check
  
  # 開始メッセージ
  print("Excel-CSV変換処理を開始します。")
  print("入力ファイル：" + IN_FILE)

  print("...")

  # 変換処理
  outFiles = Excel2Csv(IN_FILE, OUT_PATH)
  if(outFile is None):
    return 99

  # 終了メッセージ
  print("Excel-CSV変換処理が完了しました。\n出力ファイル：")
  for outFile in outFiles:
    print("・" + outFile)

  return 0


# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
# % 関数定義 %
# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
main(sys.argv)
