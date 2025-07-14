#%%
import re

import pandas as pd

from mail_handler import sending_email
from model import *

gc = authorize_gspread()
spreadsheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1Jjywp-EBUkG_ICkD4PXM9_Tp1E4g1Pigbm4sJxxVHoU/edit?gid=1368181872#gid=1368181872')
#%%
# 選擇第一個工作表
all_res = []
#%%
for ws in spreadsheet.worksheets():
    all_records = ws.get_all_records()
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

df['時間戳記'] = pd.to_datetime(df['時間戳記'].str.split(' ').str[0], format='%Y/%m/%d', errors='coerce')
df = df[df['時間戳記'] >= pd.Timestamp('2025-07-07')]
df['時間戳記'] = df['時間戳記'].astype(str)
# 以電子郵件分組，彙整所有今日問題
result = df.groupby('電子郵件地址').agg({
    '今日問題彙整': lambda x: "\n".join(x),
    '時間戳記': lambda x: ", ".join(x)
}).reset_index()
# 查看結果
print(result)
result.values.tolist()
#%%

result_index = -2
text = result.values.tolist()[result_index][1]
date_list = [t.split()[0] for t in result.values.tolist()[result_index][2].split(', ')]
content_list = [f'<p style="font-size:16px; color:#4c70a0; font-weight:bold;">{line.strip()}</p>' for line in text.split('\n今日') if line.strip()]

for i in range(len(content_list)):
    qmatch = re.search(r'：(.*?)？：', content_list[i]).group(1)
    content_list[i] = re.sub(
        r'(今日)?問題：',
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
<h2 style="text-align:center;">喜歡一天當中的什麼時候？</h2>
<p style="font-size:16px; text-align:justify;">要問我一天當中喜歡什麼時候？大約是清晨，雖然我並非是一個習慣早起的人，只是覺得，在那個時段才能獲得片刻的庇護。</p>
<p style="font-size:16px; text-align:justify;">當村上春樹說「世界在清晨甦醒過來，帶著一種沈默的呼吸。」這句充滿生命力卻又不喧囂的氛圍，給人沉靜而有力的感受。不像白天的喧囂，所有的人事物都爭先恐後地叫囂著，讓人只想把耳朵摀起來，甚至把頭埋進沙裡。</p>
<p style="font-size:16px; text-align:justify;">我最喜歡的，莫過於清晨房間裡透出微光的那種靜謐。光線是柔和的，還帶著一絲涼意，像極了這個世界難得的溫柔。空氣中還瀰漫著一絲夜晚的餘溫，沒有白天那種刺眼的、讓人只想逃離的壓迫感，給人一種平靜的氛圍。</p>
<p style="font-size:16px; text-align:justify;">如果身邊剛好有朋友或伴侶睡在旁邊，那更是種奢侈的享受，不用擔心誰會打擾，也不用擔心說出口的話會被誰批判。我們可以毫無顧忌地聊著今天能做些什麼，也許是去巷口那家新開的咖啡店，也許是窩在家裡看一部老電影，或者只是單純地聊著那些稀奇古怪的問題：結婚前要不要簽離婚協議書？然後，聊著聊著，就再次沉沉睡去。</p>
<p style="font-size:16px; text-align:justify;">這段音樂旅程，從獨自觀賞演唱會的忐忑，到勇敢跨足海外獨旅，再到雨中音樂祭的體驗，每一次都是對自我的探索與突破，擁抱未知的可能性，也更懂得在不同情境中尋找屬於自己的快樂。</p>
<p style="font-size:16px; text-align:justify;">我喜歡清晨，不是因為它多麼充滿希望，而是因為它充滿了「空白」。那空白允許我們暫時放下所有的疲憊和不滿，允許我們在世界真正開始喧囂之前，偷偷喘口氣，做一場只有自己才懂的白日夢。清晨的光，像一雙溫柔的手，輕輕撫慰著被現實折磨得千瘡百孔的靈魂。在那片刻的寧靜中，我們得以充電、喘息，才有力氣繼續面對接下來一整天無休止的喧囂。</p>
<p style="font-size:16px; text-align:justify;">你呢？你喜歡一天當中的哪個時刻？</p>
<p style="font-size:16px; text-align:justify;">如果想要一個人走走晃晃，有own的7月展覽推薦：<br>https://www.instagram.com/p/DLpRU_QTstr/</p>
<p style="font-size:16px; text-align:justify;">📢own即將於9月推出《2026 own Question Diary | 指引者手帳》，陪伴你進行更深層的自我探索，與自己好好相處。</p>
<p style="font-size:16px; text-align:justify;">歡迎持續關注<a href="https://www.instagram.com/own.bimonthly/">instagram</a>，才不會錯過最新消息✨</p>
'''

past_q = '''
<h2 style="text-align:center";>過往靈魂提問</h2>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/Wnc4TEvRe445Gtmr6">靈魂提問週報Q1</a></p>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/1viErqq1jFdk6DJr9">靈魂提問週報Q2</a></p>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/U9kppXKN2NoHP9u3A">靈魂提問週報Q3</a></p>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/21D7MieauK8pgZ5h7">靈魂提問週報Q4</a></p>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/UpLLbgZBa2NS16v48">靈魂提問週報Q5</a></p>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/y99vJJw7zWojWyGf7">靈魂提問週報Q6</a></p>
<p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/TG4acrAe47GWGWaG8">靈魂提問週報Q7</a></p>
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
subject = '【靈魂提問週報回顧】你喜歡一天當中的什麼時候？'
sending_email('dogsen1999@gmail.com', subject, final_html)
sending_email('ujs7171997@gmail.com', subject, final_html)
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
            r'(今日)?問題：',
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
    <h2 style="text-align:center;">喜歡一天當中的什麼時候？</h2>
    <p style="font-size:16px; text-align:justify;">要問我一天當中喜歡什麼時候？大約是清晨，雖然我並非是一個習慣早起的人，只是覺得，在那個時段才能獲得片刻的庇護。</p>
    <p style="font-size:16px; text-align:justify;">當村上春樹說「世界在清晨甦醒過來，帶著一種沈默的呼吸。」這句充滿生命力卻又不喧囂的氛圍，給人沉靜而有力的感受。不像白天的喧囂，所有的人事物都爭先恐後地叫囂著，讓人只想把耳朵摀起來，甚至把頭埋進沙裡。</p>
    <p style="font-size:16px; text-align:justify;">我最喜歡的，莫過於清晨房間裡透出微光的那種靜謐。光線是柔和的，還帶著一絲涼意，像極了這個世界難得的溫柔。空氣中還瀰漫著一絲夜晚的餘溫，沒有白天那種刺眼的、讓人只想逃離的壓迫感，給人一種平靜的氛圍。</p>
    <p style="font-size:16px; text-align:justify;">如果身邊剛好有朋友或伴侶睡在旁邊，那更是種奢侈的享受，不用擔心誰會打擾，也不用擔心說出口的話會被誰批判。我們可以毫無顧忌地聊著今天能做些什麼，也許是去巷口那家新開的咖啡店，也許是窩在家裡看一部老電影，或者只是單純地聊著那些稀奇古怪的問題：結婚前要不要簽離婚協議書？然後，聊著聊著，就再次沉沉睡去。</p>
    <p style="font-size:16px; text-align:justify;">這段音樂旅程，從獨自觀賞演唱會的忐忑，到勇敢跨足海外獨旅，再到雨中音樂祭的體驗，每一次都是對自我的探索與突破，擁抱未知的可能性，也更懂得在不同情境中尋找屬於自己的快樂。</p>
    <p style="font-size:16px; text-align:justify;">我喜歡清晨，不是因為它多麼充滿希望，而是因為它充滿了「空白」。那空白允許我們暫時放下所有的疲憊和不滿，允許我們在世界真正開始喧囂之前，偷偷喘口氣，做一場只有自己才懂的白日夢。清晨的光，像一雙溫柔的手，輕輕撫慰著被現實折磨得千瘡百孔的靈魂。在那片刻的寧靜中，我們得以充電、喘息，才有力氣繼續面對接下來一整天無休止的喧囂。</p>
    <p style="font-size:16px; text-align:justify;">你呢？你喜歡一天當中的哪個時刻？</p>
    <p style="font-size:16px; text-align:justify;">如果想要一個人走走晃晃，有own的7月展覽推薦：<br>https://www.instagram.com/p/DLpRU_QTstr/</p>
    <p style="font-size:16px; text-align:justify;">📢own即將於9月推出《2026 own Question Diary | 指引者手帳》，陪伴你進行更深層的自我探索，與自己好好相處。</p>
    <p style="font-size:16px; text-align:justify;">歡迎持續關注<a href="https://www.instagram.com/own.bimonthly/">instagram</a>，才不會錯過最新消息✨</p>

    '''

    past_q = '''
    <h2 style="text-align:center";>過往靈魂提問</h2>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/Wnc4TEvRe445Gtmr6">靈魂提問週報Q1</a></p>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/1viErqq1jFdk6DJr9">靈魂提問週報Q2</a></p>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/U9kppXKN2NoHP9u3A">靈魂提問週報Q3</a></p>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/21D7MieauK8pgZ5h7">靈魂提問週報Q4</a></p>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/UpLLbgZBa2NS16v48">靈魂提問週報Q5</a></p>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/y99vJJw7zWojWyGf7">靈魂提問週報Q6</a></p>
    <p style="text-align:center" style="font-size:16px;"><a href="https://forms.gle/TG4acrAe47GWGWaG8">靈魂提問週報Q7</a></p>
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
    subject = f'【靈魂提問週報 W3】邀請你回顧：你的獨處與自我探索'
    sending_email(personal_gmail, subject, final_html)
# %%
