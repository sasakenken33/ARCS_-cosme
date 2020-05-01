from flask import Flask, request, redirect, url_for, render_template, flash, session, send_file
from bs4 import BeautifulSoup
from openpyxl import Workbook
import requests
import os
import re

app = Flask(__name__)
app.secret_key = os.urandom(24)

XLSX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

@app.route("/", methods=["GET"])
def top():
  return render_template("top.html")

@app.route("/input", methods=["GET","POST"])
def input():
  if request.form["item_url"]:
    fetch()
    flash("レビュー収集が完了しました", "success")
    return render_template("top.html")
  else:
    flash("URLは必須です", "alert")
    return render_template("top.html")

def is_int(i):
  try:
    int(i)
    return True
  except ValueError:
    return False

def fetch():
  input_url = request.form["item_url"]

  #URL形式が正しいか確認
  if re.search("https://www.cosme.net/product/product_id/\d{6,8}/top", input_url) == None:
    flash("URLの形式が違います","alert")
    return render_template("top.html")
  
  #レビュー一覧ページのURLに変換
  rv_url = input_url.replace("top", "reviews")

  r = requests.get(rv_url)
  html = r.text
  soup = BeautifulSoup(html, "html.parser")

  #レビュー数取得
  rv_count = soup.find("span", {"class":"count cnt"}).text
   
  #レビュー数の指定はあるか
  if request.form["input_count"]:

    #レビュー数の指定がある→取得レビュー数にはintが入っているか確認
    if is_int(request.form["input_count"]):
      pass
    else:
      flash("レビュー数には半角数字を入力してください","alert")
      return render_template("top.html")  
    
    #取得レビュー数が適切かチェック
    if int(request.form["input_count"]) > int(rv_count) :
      flash("収集できるレビューは{}件までです。".format(rv_count),"alert")
      return render_template("top.html")

    #取得レビュー数を上書き
    rv_count = request.form["input_count"]
    roop_count = int(rv_count)
  #レビュー数の指定がない→掲載されている全てのレビューを取得する
  else:
    roop_count = int(rv_count)
    
  #商品名取得(出力ファイル名になる)
  item_name = soup.find("strong", {"class":"pdct-name fn"})
  item_name = item_name.text.strip()

  #全文レビューページURL取得
  items = soup.select("span > a.cmn-viewmore")
  input_url = items[0].get("href")

  #カウンター変数
  i = 0

  data_set = []

  while i < roop_count:
    r = requests.get(input_url)
    html = r.text
    soup = BeautifulSoup(html, "html.parser")
    #投稿本文
    content = soup.find("p", {"class":"read"})
    #星の数
    star = soup.find("p", {"class":"reviewer-rating"})
    #ユーザー名
    usr_name = soup.find("span", {"class":"reviewer-name"})
    #次の全文レビューページURL
    next_url = soup.select("li.next > a")
    try:
      input_url = next_url[0].get("href")
    except IndexError:
      print("完了")
    #[ユーザー名、星の数、投稿本文]の配列を追加
    usr_set = [usr_name.text, star.text, content.text]
    data_set.append(usr_set)
    #カウンター変数加算
    i+=1

  #エクセルファイル開く
  book = Workbook()
  sheet = book.active
  #一行ごとに追記
  for data in data_set:
    sheet.append(data)
  #商品名をファイル名にして保存＆ダウンロード
  file_name = "レビュー収集結果_{}.xlsx".format(item_name)
  book.save(file_name)
  session["dlfile_name"] = file_name

@app.route('/download', methods=['GET'])
def download():
  return send_file(session["dlfile_name"], as_attachment = True, attachment_filename = session["dlfile_name"], mimetype = XLSX_MIMETYPE)