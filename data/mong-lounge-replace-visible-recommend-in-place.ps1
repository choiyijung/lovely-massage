$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$index = Join-Path $root "index.html"
$js = Join-Path $root "assets\js\main-recommend-final.js"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-main-visible-replace-backup-$stamp"

if (!(Test-Path -LiteralPath $js)) { throw "파일 없음: $js" }
if (!(Test-Path -LiteralPath $index)) { throw "파일 없음: $index" }

New-Item -ItemType Directory -Force -Path $backup | Out-Null
Copy-Item -LiteralPath $js -Destination (Join-Path $backup "main-recommend-final.js") -Force
Copy-Item -LiteralPath $index -Destination (Join-Path $backup "index.html") -Force

$src = Get-Content -LiteralPath $js -Raw -Encoding UTF8

# 이전에 같은 보정을 실행했다면 제거 후 다시 추가
$src = [regex]::Replace(
    $src,
    '(?is)/\*\s*ML-VISIBLE-RECOMMEND-REPLACE-START\s*\*/.*?/\*\s*ML-VISIBLE-RECOMMEND-REPLACE-END\s*\*/',
    ''
)

$patch = @'

/* ML-VISIBLE-RECOMMEND-REPLACE-START */
(function(){
  function esc(v){
    return String(v == null ? "" : v)
      .replace(/&/g,"&amp;")
      .replace(/</g,"&lt;")
      .replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;");
  }

  function getData(){
    return window.ML_MAIN_RECOMMEND ||
           window.ML_MAIN_RECOMMEND_FINAL ||
           {vip:[], premium:[]};
  }

  function findHeading(title){
    return Array.from(document.querySelectorAll("h1,h2,h3,h4,strong,b"))
      .find(el=>{
        if(el.closest("#ml-vip-recommend-live,#ml-premium-recommend-live")) return false;
        return (el.textContent || "").replace(/\s+/g," ").trim() === title;
      }) || null;
  }

  function findVisibleRecommendBlock(title){
    const h = findHeading(title);
    if(!h) return null;

    let n = h;
    let best = null;

    for(let i=0; i<8 && n && n !== document.body; i++, n=n.parentElement){
      const text = (n.innerText || "").replace(/\s+/g," ");
      const details = (text.match(/상세\s*보기/g) || []).length;

      if(details >= 4){
        best = n;
        break;
      }

      if(n.tagName === "MAIN") break;
    }
    return best;
  }

  function card(x,badge){
    const href = x.link && x.link !== "#" ? x.link : "javascript:void(0)";
    return `
      <a class="ml-live-card" href="${esc(href)}">
        <div class="ml-live-thumb">
          <img src="${esc(x.image || "")}" alt="${esc(x.shop || "추천 업체")}">
          <span>${esc(badge)}</span>
          <b>${esc(x.region || "추천 지역")}</b>
        </div>
        <div class="ml-live-body">
          <h3>${esc(x.shop || "업체")}</h3>
          <p>${esc(x.desc || "상세페이지에서 이용 정보를 확인해보세요.")}</p>
          <div class="ml-live-foot">
            <strong>${esc(x.price || "가격 문의")}</strong>
            <em>상세 보기</em>
          </div>
        </div>
      </a>`;
  }

  function makeSection(id,title,badge,items){
    const sec = document.createElement("section");
    sec.id = id;
    sec.className = "ml-live-section";
    sec.innerHTML = `
      <div class="ml-live-head">
        <h2>${esc(title)}</h2>
        <span>전체보기</span>
      </div>
      <div class="ml-live-grid">${(items || []).slice(0,4).map(x=>card(x,badge)).join("")}</div>`;
    return sec;
  }

  function ensureStyle(){
    if(document.getElementById("ml-live-recommend-style")) return;

    const s = document.createElement("style");
    s.id = "ml-live-recommend-style";
    s.textContent = `
      .ml-live-section{margin:30px 0 34px!important;padding:0!important}
      .ml-live-head{display:flex!important;align-items:center!important;justify-content:space-between!important;margin:0 0 13px!important}
      .ml-live-head h2{margin:0!important;font-size:22px!important;color:#111!important}
      .ml-live-head span{font-size:11px!important;color:#777!important}
      .ml-live-grid{display:grid!important;grid-template-columns:repeat(4,minmax(0,1fr))!important;gap:10px!important}
      .ml-live-card{display:block!important;overflow:hidden!important;border:1px solid #ddd!important;border-radius:14px!important;background:#fff!important;color:#111!important;text-decoration:none!important;box-shadow:0 2px 8px rgba(0,0,0,.05)!important}
      .ml-live-thumb{height:120px!important;position:relative!important;overflow:hidden!important;background:#174f43!important}
      .ml-live-thumb img{width:100%!important;height:100%!important;object-fit:cover!important;display:block!important}
      .ml-live-thumb span{position:absolute!important;left:0!important;top:0!important;background:#efc342!important;color:#111!important;padding:14px 20px!important;border-radius:0 0 14px 0!important;font-size:11px!important;font-weight:900!important}
      .ml-live-thumb b{position:absolute!important;left:12px!important;bottom:10px!important;color:#fff!important;font-size:12px!important;text-shadow:0 1px 3px #000!important}
      .ml-live-body{padding:12px!important}
      .ml-live-body h3{margin:0 0 6px!important;font-size:15px!important}
      .ml-live-body p{height:38px!important;overflow:hidden!important;margin:0 0 14px!important;color:#777!important;font-size:11px!important;line-height:1.55!important}
      .ml-live-foot{display:flex!important;align-items:center!important;justify-content:space-between!important;gap:8px!important}
      .ml-live-foot strong{font-size:14px!important}
      .ml-live-foot em{font-style:normal!important;border:1px solid #e5ad32!important;border-radius:8px!important;padding:6px 9px!important;font-size:10px!important;font-weight:800!important}
      @media(max-width:760px){.ml-live-grid{grid-template-columns:repeat(2,minmax(0,1fr))!important}}
    `;
    document.head.appendChild(s);
  }

  function replaceOne(title,id,badge,items){
    if(document.getElementById(id)) return true;
    if(!items || !items.length) return false;

    const old = findVisibleRecommendBlock(title);
    if(!old) return false;

    const sec = makeSection(id,title,badge,items);
    old.replaceWith(sec);
    return true;
  }

  function apply(){
    ensureStyle();

    // 이전 스크립트가 아래쪽에 추가한 별도 추천영역은 제거
    const extra = document.getElementById("ml-main-recommend-final");
    if(extra) extra.remove();

    const d = getData();
    replaceOne("VIP 추천","ml-vip-recommend-live","VIP",d.vip || []);
    replaceOne("프리미엄 추천","ml-premium-recommend-live","PICK",d.premium || []);
  }

  function boot(){
    apply();
    requestAnimationFrame(apply);
    setTimeout(apply,100);
    setTimeout(apply,400);
    setTimeout(apply,1000);
    setTimeout(apply,2000);
  }

  if(document.readyState === "loading"){
    document.addEventListener("DOMContentLoaded",boot);
  } else {
    boot();
  }
})();
/* ML-VISIBLE-RECOMMEND-REPLACE-END */
'@

$src = $src.TrimEnd() + "`r`n" + $patch + "`r`n"
Set-Content -LiteralPath $js -Value $src -Encoding UTF8

# index의 final.js 버전을 강제로 새 값으로 변경
$html = Get-Content -LiteralPath $index -Raw -Encoding UTF8
$html = [regex]::Replace(
    $html,
    'assets/js/main-recommend-final\.js(?:\?v=[^"'']*)?',
    "assets/js/main-recommend-final.js?v=$stamp"
)
Set-Content -LiteralPath $index -Value $html -Encoding UTF8

$verifyJs = Get-Content -LiteralPath $js -Raw -Encoding UTF8
$verifyIndex = Get-Content -LiteralPath $index -Raw -Encoding UTF8

Write-Host ""
Write-Host "=== 메인 보이는 추천카드 직접 교체 보정 완료 ==="
Write-Host "기존 VIP 자리 직접 교체 코드:" ($verifyJs -match 'ml-vip-recommend-live')
Write-Host "기존 프리미엄 자리 직접 교체 코드:" ($verifyJs -match 'ml-premium-recommend-live')
Write-Host "별도 추천영역 제거 코드:" ($verifyJs -match 'ml-main-recommend-final')
Write-Host "캐시 버전 갱신:" ($verifyIndex -match [regex]::Escape("main-recommend-final.js?v=$stamp"))
Write-Host "백업 폴더:" $backup
Write-Host ""
Write-Host "이제 5512 주소에서 Ctrl+F5 하세요."
