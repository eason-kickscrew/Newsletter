#%%
import re

import pandas as pd

from mail_handler import sending_email
from model import *

gc = authorize_gspread()
spreadsheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1Jjywp-EBUkG_ICkD4PXM9_Tp1E4g1Pigbm4sJxxVHoU/edit?gid=1368181872')
#%%
# 選擇第一個工作表
all_res =[]
for i in range(1, 8):
    worksheet1 = spreadsheet.worksheet(f'表單回應 {i}')
    all_records = worksheet1.get_all_records()
    all_res.extend(all_records)
# %%
df = pd.DataFrame(all_res)

# 選出所有包含「今日問題」的欄位
today_question_cols = [col for col in df.columns if '今日問題' in col]

# 依電子郵件分組，把所有「今日問題」的 key-value 合併成一段文字
def combine_questions(row):
    texts = []
    for col in today_question_cols:
        if pd.notna(row.get(col)):
            texts.append(f"{col}：{row[col]}")
    return "\n".join(texts)

df['今日問題彙整'] = df.apply(combine_questions, axis=1)

# 以電子郵件分組，彙整所有今日問題
result = df.groupby('電子郵件地址').agg({
    '今日問題彙整': lambda x: "\n".join(x),
    '時間戳記': lambda x: ", ".join(x)
}).reset_index()
# 查看結果
print(result)
result.values.tolist()


result_index = 23
text = result.values.tolist()[result_index][1]
date_list = [t.split()[0] for t in result.values.tolist()[result_index][2].split(', ')]
content_list = [f'<p style="font-size:16px; color:#4c70a0; font-weight:bold;">{line.strip()}</p>' for line in text.split('\n今日') if line.strip()]

for i in range(len(content_list)):
    qmatch = re.search(r'：(.*?)？：', content_list[i]).group(1)
    content_list[i] = re.sub(
        r'問題：',
        f'來自 {date_list[i]} 的你：',
        content_list[i],
        count=1
    )
    content_list[i] = re.sub(
        r'：(.*?？：)',
        f'<br>「{qmatch}？」</p><p style="font-size:16px;">',
        content_list[i],
        count=1
    )
    
content_html = '<br>'.join(content_list)
content_html = '<h2 style="text-align:center"> 每週問題回顧</h2>' + content_html

banner_url = 'https://i.ibb.co/gZzDwXpm/banner.jpg'  # 圖片直連網址，Gmail 可顯示
banner_html = f'<div style="text-align:center;"><img src="{banner_url}" alt="banner" style="width:100%;"></div>'
pw_url = 'https://media.discordapp.net/attachments/1346850023601606720/1381288449298141254/2.jpg?ex=6858c4dd&is=6857735d&hm=b4b279df54c462e585c5de4f4193300d84c84ceac1f0d1deb758aa8dc1fc2fd4&=&format=webp&width=943&height=256'
pw_html = f'<div style="text-align:center;"><img src="{pw_url}" alt="banner" style="width:100%;"></div>'
yen_url = 'https://i.ibb.co/7Jvjmk09/yen.jpg'
yen_html = f'<p style="font-size:16px;">by own 募資策劃 Yen</p><div style="text-align:center;"><img src="{yen_url}" alt="banner" style="width:100%;"></div>'


feedback_html = '''
<h2 style="text-align:center;">你的理想生活是什麼狀態？</h2>
<p style="font-size:16px; text-align:justify;">這問題來自於某天週末早晨，原先預計要睡到下午來彌補一週上班的疲累，結果7點就清醒了，想著自己怎麼這麼沒用居然睡不著了！！只好認命起床。準備踏進浴室洗漱前，遇到剛好踏出房門準備要去加班的室友。</p>
<p style="font-size:16px; text-align:justify;">室友：你今天要做什麼？</p>
<p style="font-size:16px; text-align:justify;">我：哦～等等去健身房運動1小時，回家洗澡吃完早餐後就出門去逛展覽，下午去書店買書再去甜點店看書或是寫手帳吧</p>
<p style="font-size:16px; text-align:justify;">室友突然用著一種快落淚的語氣回我：這是我的理想退休生活耶，但我現在卻要去公司加班！</p>
<p style="font-size:16px; text-align:justify;">我忽然意識到，是啊，這正是我退休後想過的生活，運動讓自己情緒穩定、身體健康，閱讀讓自己透過更多人的眼睛認識世界運轉的樣子，逛展覽讓自己的生活有更多啟發，寫手帳則是一種輸入資訊後輸出和自己對話思考的過程。</p>
<p style="font-size:16px; text-align:justify;">原來我就活在我理想的生活中，找到工作與生活的平衡。先前對於平日上班下班一成不變的生活感到疲乏、不滿意，但其實好好享受每一個當下，不管是與人相處或獨處，只要是自己想要的，就是在理想生活中了。</p>
<p style="font-size:16px; text-align:justify;">如果想要一個人走走晃晃，有own的6月展覽推薦：<a href="https://www.instagram.com/p/DKZctptOIvI/">https://www.instagram.com/p/DKZctptOIvI/</a> 可以參考！</p>
'''

past_q = '''
<h2 style="text-align:center";>過往靈魂提問</h2>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/GfT4PAC2TrN52NAh9">靈魂提問週報Q1</a></p>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/rPm22gKZcpS92fSD6">靈魂提問週報Q2</a></p>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/dvbguh6Ur6NkErvB9">靈魂提問週報Q3</a></p>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/VsGE3ygJdtAwGD3NA">靈魂提問週報Q4</a></p>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/8DNdxKaCTpF5K4gd8">靈魂提問週報Q5</a></p>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/XQi3bc9AWRDqmgnd8">靈魂提問週報Q6</a></p>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/bDn38kqco7Y4a89W9">靈魂提問週報Q7</a></p>
'''

start_html = '''
<!DOCTYPE html>
<html>
  <body style="margin:0; padding:0;">
    <div style="max-width:700px; margin:auto;">
  '''
end_html = '''
    </div>
  </body>
</html>
'''
final_html = start_html + banner_html + feedback_html + yen_html + content_html + past_q + end_html
# %%
subject = '【靈魂提問週報W1】邀請你回顧：你的獨處與自我探索10'
sending_email('dogsen1999@gmail.com', subject, final_html)
# %%
for i in ['ujs7171997@gmail.com', 'pw39winnie@gmail.com', 'mn1080305@gmail.com']:
    sending_email(i, subject, final_html)
# %%












#%%
for i in range(len(result.values.tolist())):
    print(i)
    result_index = i
    text = result.values.tolist()[result_index][1]
    date_list = [t.split()[0] for t in result.values.tolist()[result_index][2].split(', ')]
    content_list = [f'<p style="font-size:16px; color:#4c70a0; font-weight:bold;">{line.strip()}</p>' for line in text.split('\n今日') if line.strip()]

    for i in range(len(content_list)):
        qmatch = re.search(r'：(.*?)？：', content_list[i]).group(1)
        content_list[i] = re.sub(
            r'問題：',
            f'來自 {date_list[i]} 的你：',
            content_list[i],
            count=1
        )
        content_list[i] = re.sub(
            r'：(.*?？：)',
            f'<br>「{qmatch}？」</p><p style="font-size:16px;">',
            content_list[i],
            count=1
        )
        
    content_html = '<br>'.join(content_list)
    content_html = '<h2 style="text-align:center"> 每週問題回顧</h2>' + content_html

    banner_url = 'https://i.ibb.co/gZzDwXpm/banner.jpg'  # 圖片直連網址，Gmail 可顯示
    banner_html = f'<div style="text-align:center;"><img src="{banner_url}" alt="banner" style="width:100%;"></div>'
    pw_url = 'https://media.discordapp.net/attachments/1346850023601606720/1381288449298141254/2.jpg?ex=6858c4dd&is=6857735d&hm=b4b279df54c462e585c5de4f4193300d84c84ceac1f0d1deb758aa8dc1fc2fd4&=&format=webp&width=943&height=256'
    pw_html = f'<div style="text-align:center;"><img src="{pw_url}" alt="banner" style="width:100%;"></div>'
    yen_url = 'https://i.ibb.co/7Jvjmk09/yen.jpg'
    yen_html = f'<p style="font-size:16px;">by own 募資策劃 Yen</p><div style="text-align:center;"><img src="{yen_url}" alt="banner" style="width:100%;"></div>'


    feedback_html = '''
    <h2 style="text-align:center;">你的理想生活是什麼狀態？</h2>
    <p style="font-size:16px; text-align:justify;">這問題來自於某天週末早晨，原先預計要睡到下午來彌補一週上班的疲累，結果7點就清醒了，想著自己怎麼這麼沒用居然睡不著了！！只好認命起床。準備踏進浴室洗漱前，遇到剛好踏出房門準備要去加班的室友。</p>
    <p style="font-size:16px; text-align:justify;">室友：你今天要做什麼？</p>
    <p style="font-size:16px; text-align:justify;">我：哦～等等去健身房運動1小時，回家洗澡吃完早餐後就出門去逛展覽，下午去書店買書再去甜點店看書或是寫手帳吧</p>
    <p style="font-size:16px; text-align:justify;">室友突然用著一種快落淚的語氣回我：這是我的理想退休生活耶，但我現在卻要去公司加班！</p>
    <p style="font-size:16px; text-align:justify;">我忽然意識到，是啊，這正是我退休後想過的生活，運動讓自己情緒穩定、身體健康，閱讀讓自己透過更多人的眼睛認識世界運轉的樣子，逛展覽讓自己的生活有更多啟發，寫手帳則是一種輸入資訊後輸出和自己對話思考的過程。</p>
    <p style="font-size:16px; text-align:justify;">原來我就活在我理想的生活中，找到工作與生活的平衡。先前對於平日上班下班一成不變的生活感到疲乏、不滿意，但其實好好享受每一個當下，不管是與人相處或獨處，只要是自己想要的，就是在理想生活中了。</p>
    <p style="font-size:16px; text-align:justify;">如果想要一個人走走晃晃，有own的6月展覽推薦：<a href="https://www.instagram.com/p/DKZctptOIvI/">https://www.instagram.com/p/DKZctptOIvI/</a> 可以參考！</p>
    '''

    past_q = '''
    <h2 style="text-align:center";>過往靈魂提問</h2>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/GfT4PAC2TrN52NAh9">靈魂提問週報Q1</a></p>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/rPm22gKZcpS92fSD6">靈魂提問週報Q2</a></p>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/dvbguh6Ur6NkErvB9">靈魂提問週報Q3</a></p>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/VsGE3ygJdtAwGD3NA">靈魂提問週報Q4</a></p>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/8DNdxKaCTpF5K4gd8">靈魂提問週報Q5</a></p>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/XQi3bc9AWRDqmgnd8">靈魂提問週報Q6</a></p>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/bDn38kqco7Y4a89W9">靈魂提問週報Q7</a></p>
    '''

    start_html = '''
    <!DOCTYPE html>
    <html>
    <body style="margin:0; padding:0;">
        <div style="max-width:700px; margin:auto;">
    '''
    end_html = '''
        </div>
    </body>
    </html>
    '''
    final_html = start_html + banner_html + feedback_html + yen_html + content_html + past_q + end_html
    personal_gmail = result.values.tolist()[result_index][0]
    subject = f'【靈魂提問週報 W1】邀請你回顧：你的獨處與自我探索'
    sending_email(personal_gmail, subject, final_html)
# %%
