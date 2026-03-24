import asyncio
from playwright.async_api import async_playwright
import pandas as pd
import re
import plotly.graph_objects as go

async def scrape_to_github_pages():
    async with async_playwright() as p:
        # 啟動瀏覽器
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={'width': 1920, 'height': 1080}
        )
        page = await context.new_page()

        print("🚀 啟動 MIS 雲端抓取引擎...")
        await page.goto("https://busadm.ccu.edu.tw/p/412-1248-3236.php?Lang=zh-tw", wait_until="networkidle")

        # 1. 精準識別 21 位專任教師名單
        links = await page.query_selector_all("a")
        temp_list = []
        blacklist = ["企研所", "行銷所", "博士班", "系主任", "系辦公室", "教師介紹"]
        for link in links:
            name = (await link.inner_text()).strip()
            href = await link.get_attribute("href")
            if href and "/p/4" in href and 2 <= len(name) <= 3 and name not in blacklist:
                url = href if href.startswith("http") else "https://busadm.ccu.edu.tw" + href
                if not any(t['姓名'] == name for t in temp_list):
                    temp_list.append({"姓名": name, "連結": url})

        # 鎖定範圍：從黃正魁到陳維婷
        try:
            names_only = [t['姓名'] for t in temp_list]
            final_list = temp_list[names_only.index("黃正魁"):names_only.index("陳維婷")+1]
        except:
            final_list = temp_list

        print(f"✅ 識別成功：共 {len(final_list)} 位專任教師。開始深度擷取個人頁面資料...")

        results = []
        for t in final_list:
            try:
                print(f"📡 擷取: {t['姓名']}...")
                await page.goto(t['連結'], wait_until="domcontentloaded", timeout=45000)

                # 許嘉文老師特別加強等待
                if t['姓名'] == "許嘉文":
                    await asyncio.sleep(6)
                else:
                    await asyncio.sleep(2.5)

                body_text = await page.inner_text("body")

                # A. 瀏覽數
                views_match = re.search(r"瀏覽數[:：]\s*(\d+[,]?\d*)", body_text)
                views = int(views_match.group(1).replace(',', '')) if views_match else 0

                # B. 職稱識別
                title = "專任教師"
                if "助理教授" in body_text: title = "助理教授"
                elif "副教授" in body_text: title = "副教授"
                elif "教授" in body_text: title = "教授"

                # 行政兼職硬編碼 (Exception Handling)
                if t['姓名'] == "黃正魁": title = "教授兼企管系主任"
                elif t['姓名'] == "連雅惠": title = "教授兼管理學院院長"
                elif t['姓名'] == "盧龍泉": title = "教授兼學務長"

                # C. 研究室修正
                if t['姓名'] == "盧龍泉": room = "436 ; 010"
                elif t['姓名'] in ["賴璽方", "鍾憲瑞"]: room = "洽系辦"
                elif t['姓名'] == "劉敏熙": room = "434室"
                else:
                    room_match = re.search(r"研究室[:：]?\s*(\d{3,4})", body_text)
                    room = f"{room_match.group(1)}室" if room_match else "院內研究室"

                # D. 校內分機修正
                if t['姓名'] == "盧龍泉": extension = "34311 ; 12100"
                elif t['姓名'] == "鍾憲瑞": extension = "洽系辦"
                elif t['姓名'] == "劉敏熙": extension = "34322"
                else:
                    ext_internal = re.search(r"校內[:：]\s*(\d{5})", body_text)
                    extension = ext_internal.group(1) if ext_internal else "34319"

                # E. Email 修正
                email_special = {
                    "曾光華": "marketingkfc@gmail.com",
                    "莊世杰": "bmascc@gmail.com",
                    "鍾憲瑞": "hj0730.chung@gmail.com",
                    "蘇宏仁": "bmaesu@gmail.com",
                    "陳明德": "ft.takasi@gmail.com"
                }
                
                if t['姓名'] in email_special:
                    email = email_special[t['姓名']]
                else:
                    email_match = re.search(r"([a-zA-Z0-9._-]+@ccu\.edu\.tw)", body_text)
                    email = email_match.group(1) if email_match else "未標註"

                # 建立超連結
                link_text = f'<a href="{t["連結"]}" target="_blank" style="color:#00FFCC;text-decoration:none;">點擊前往</a>'

                results.append({
                    "姓名": t['姓名'], "職稱": title, "研究室": room, "分機": extension,
                    "電子郵件": email, "瀏覽數": views, "詳細資料": link_text,
                    "原始連結": t['連結']
                })
            except Exception as e:
                print(f"⚠️ {t['姓名']} 抓取失敗: {e}")

        await browser.close()

        # --- 2. 數據處理與產出網頁 ---
        if results:
            df = pd.DataFrame(results).sort_values(by="瀏覽數", ascending=False)
            df['排名'] = range(1, len(df) + 1)

            # 在 Log 紀錄中印出資料表格 (方便偵錯)
            print("\n📊 抓取結果摘要：")
            print(df[['排名', '姓名', '職稱', '瀏覽數']].to_string(index=False))

            # --- 3. 建立 Plotly 互動圖表 ---
            df_plot = df.iloc[::-1].copy()
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=df_plot['瀏覽數'],
                y=df_plot['姓名'],
                orientation='h',
                customdata=df_plot[['排名', '姓名', '職稱', '研究室', '分機', '電子郵件', '原始連結']],
                marker=dict(color=df_plot['瀏覽數'], colorscale='Sunsetdark', line=dict(color='white', width=1)),
                hovertemplate=(
                    "<b>排名: 第 %{customdata[0]} 名</b><br>" +
                    "姓名: %{customdata[1]}<br>" +
                    "職稱: %{customdata[2]}<br>" +
                    "研究室: %{customdata[3]}<br>" +
                    "分機: %{customdata[4]}<br>" +
                    "Email: %{customdata[5]}<br>" +
                    "瀏覽數: %{x:,}<br>" +
                    "<i>點擊前往介紹頁</i><extra></extra>"
                )
            ))

            fig.update_layout(
                title=dict(text="國立中正大學企管系 - 21位專任教師網頁熱度即時更新", font=dict(size=26, color='#00FFCC')),
                template="plotly_dark",
                xaxis_title="網頁總瀏覽次數 (Page Views)",
                xaxis=dict(tickformat=","),
                yaxis=dict(tickfont=dict(size=13)),
                height=1000,
                margin=dict(l=120, r=60, t=110, b=60)
            )

            # 4. 存檔為 index.html (供 GitHub Pages 使用)
            html_content = fig.to_html(include_plotlyjs='cdn', config={'displayModeBar': False})
            link_script = """
            <script>
                var plot_el = document.getElementsByClassName('plotly-graph-div')[0];
                plot_el.on('plotly_click', function(data){
                    var url = data.points[0].customdata[6];
                    if(url) window.open(url, '_blank');
                });
            </script>
            """
            with open("index.html", "w", encoding="utf-8") as f:
                f.write(html_content + link_script)
            
            print("✅ index.html 檔案產出成功！")
        else:
            print("❌ 失敗：沒有抓到任何資料。")

if __name__ == "__main__":
    asyncio.run(scrape_to_github_pages())
