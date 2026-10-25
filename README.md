# 연맹 인증 봇 설치
1. https://discord.com/developers/applications → New Application → Bot → 토큰 복사, **Server Members Intent** 켜기
2. OAuth2 → URL Generator: scope `bot`, `applications.commands` / 권한 `Manage Roles`, `Manage Nicknames`, `Send Messages` → 서버에 초대
3. 서버에 역할 `동맹원`, `간부` 생성. 서버 설정 → 역할에서 **봇 역할을 두 역할보다 위로** 올리기
4. `.env.example`을 `.env`로 복사해 값 입력, 엑셀(261023.xlsx)을 같은 폴더에 두기
5. `pip install -r requirements.txt` → `python bot.py` (PC가 켜져 있는 동안만 동작. 상시 운영은 VPS 권장)
6. 인증 채널에서 `/인증패널` 실행

## 상황판(웹)
- 봇이 켜져 있으면 http://localhost:8080/ 에서 인사 상황판이 열립니다(.env의 WEB_PORT).
- `/명단업데이트`를 실행하면 시트의 인사·톡방참여자 탭을 다시 읽어 상황판이 바로 바뀝니다.
- 밖에서 접속하려면 터널이 필요합니다. 예: `cloudflared tunnel --url http://localhost:8080`
  → 나오는 https://....trycloudflare.com 주소를 공유합니다(실행할 때마다 바뀜). 고정 주소는 서버(VPS) 사용.
- WEB_KEY를 설정하면 주소 뒤에 ?key=값 이 있어야 열립니다.
