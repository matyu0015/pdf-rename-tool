#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
日程表PDFのテストファイル作成スクリプト
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# 日本語フォントの登録（macOSの場合）
try:
    pdfmetrics.registerFont(TTFont('Japanese', '/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc'))
    font_name = 'Japanese'
except:
    try:
        pdfmetrics.registerFont(TTFont('Japanese', '/Library/Fonts/Arial Unicode.ttf'))
        font_name = 'Japanese'
    except:
        font_name = 'Helvetica'

def create_schedule_pdf():
    """日程表PDFを作成"""

    # PDFファイル名
    filename = "test_schedule.pdf"

    # PDFドキュメントの作成
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        rightMargin=20*mm,
        leftMargin=20*mm,
        topMargin=20*mm,
        bottomMargin=20*mm
    )

    # ストーリー（コンテンツ）のリスト
    story = []

    # スタイルの設定
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=16,
        textColor=colors.HexColor('#1a56db'),
        alignment=TA_CENTER,
        spaceAfter=20
    )

    # タイトル
    title = Paragraph("面接日程表", title_style)
    story.append(title)
    story.append(Spacer(1, 10*mm))

    # テーブルデータの作成
    data = [
        ['日付', '時間', '会場名', '面接官', '備考'],
    ]

    # 3月12日のデータ
    date_3_12 = ['3月12日', '', '', '', '']
    data.append(date_3_12)
    data.append(['', '10:00', 'オンライン', '近藤', ''])
    data.append(['', '11:00', 'オンライン', '近藤', ''])
    data.append(['', '13:00', 'オンライン', '近藤', ''])
    data.append(['', '15:00', '', '', ''])
    data.append(['', '', '', '', ''])  # 空行

    # 3月15日のデータ
    date_3_15 = ['3月15日', '', '', '', '']
    data.append(date_3_15)
    data.append(['', '09:00', 'オンライン', '田中', ''])
    data.append(['', '14:00', 'オンライン', '田中', ''])
    data.append(['', '16:00', 'オンライン', '田中', ''])
    data.append(['', '', '', '', ''])  # 空行

    # 4月10日のデータ
    date_4_10 = ['4月10日', '', '', '', '']
    data.append(date_4_10)
    data.append(['', '10:00', '本社会議室', '佐藤', ''])
    data.append(['', '11:30', '本社会議室', '佐藤', ''])
    data.append(['', '14:00', '本社会議室', '佐藤', ''])
    data.append(['', '15:30', '本社会議室', '佐藤', ''])

    # テーブルの作成
    table = Table(data, colWidths=[30*mm, 20*mm, 35*mm, 25*mm, 40*mm])

    # テーブルスタイルの設定
    table.setStyle(TableStyle([
        # ヘッダー行のスタイル
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#eff6ff')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#1a56db')),
        ('FONTNAME', (0, 0), (-1, 0), font_name),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('FONTNAME', (0, 1), (-1, -1), font_name),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),

        # 全体のグリッド
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
        ('BOX', (0, 0), (-1, -1), 2, colors.HexColor('#1a56db')),

        # 日付セルの背景色
        ('BACKGROUND', (0, 1), (0, 1), colors.HexColor('#f0f2ff')),  # 3月12日
        ('BACKGROUND', (0, 7), (0, 7), colors.HexColor('#f0f2ff')),  # 3月15日
        ('BACKGROUND', (0, 11), (0, 11), colors.HexColor('#f0f2ff')), # 4月10日

        # 日付セルのフォント
        ('FONTNAME', (0, 1), (0, 1), font_name),
        ('FONTSIZE', (0, 1), (0, 1), 12),
        ('FONTNAME', (0, 7), (0, 7), font_name),
        ('FONTSIZE', (0, 7), (0, 7), 12),
        ('FONTNAME', (0, 11), (0, 11), font_name),
        ('FONTSIZE', (0, 11), (0, 11), 12),

        # 行の高さ調整
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')]),
    ]))

    story.append(table)

    # PDFの生成
    doc.build(story)
    print(f"✓ テスト用PDFを作成しました: {filename}")
    return filename

if __name__ == '__main__':
    create_schedule_pdf()
