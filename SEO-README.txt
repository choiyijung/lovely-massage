전국테라피 14차 SEO 정리본

현재 적용:
- 전체 HTML canonical 추가 (상대 canonical: 현재 배포 도메인 기준 자동 해석)
- index/follow 및 noindex/follow 분리
- 주요 페이지 title/meta description 정리
- OG title/description/type 동기화
- robots.txt 생성
- sitemap 대상 경로 목록 생성

중요:
sitemap.xml은 절대 URL이 필수이므로 실제 도메인을 임의로 만들지 않았습니다.
도메인이 정해진 뒤 PowerShell에서 아래처럼 실행하면 sitemap.xml과 robots.txt의 Sitemap 줄이 자동 생성됩니다.

예:
powershell -ExecutionPolicy Bypass -File .\build-sitemap.ps1 -Domain "https://example.com"

noindex 처리:
- /search.html
- /login.html
- /join.html
- /mypage.html
- /shop.html

독립 지역 페이지와 정적 업체 상세페이지는 index/follow입니다.
