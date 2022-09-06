# 키매크로, 웹컨트롤 기능이 있는 액셀 매크로
## 1. 일일 업무보고 형식의 예제
다음주 날짜 추가, 업무 항목 추가 기능 포함.   
웹컨트롤를 이용한 자동 업로드 (hiwork site 기준으로 작성)   
다음주 시트 작성시 추가 되지 않음. (이번주만 시트 생성 함.)
## 2. 사용법
다른 부분 수정 없이 Sub web_control() 부분만 수정하면 됩니다.
```vba
[line 68]
URL = "https://office.hiworks.com/your_domain/home/logout" '로그아웃 URL

[line 72]
URL = "https://office.hiworks.com/your_domain/bbs/board/board_write/modify/4321/123" '본인 게시물 주소

[line 84]
Set input_Data = HTMLDoc.getElementById("office_id")
input_Data.Value = "Your ID" 'ID 입력
    
Set input_Data = HTMLDoc.getElementById("office_passwd")
input_Data.Value = "Your PW" '암호 입력

[line 102]
'파일 경로 수정
Sleep (1100)
setClip ("D:\Documents\일일업무보고_양식.xlsm")
```