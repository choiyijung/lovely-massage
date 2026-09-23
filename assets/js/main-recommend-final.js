window.ML_MAIN_RECOMMEND = {"vip":[{"shop":"한국미인테라피","region":"서울 · 경기 · 인천","price":"100,000원~","image":"/assets/images/main-recommend/vip-01.svg","link":"/gangnam-gu/shops/001/","desc":"한국미인테라피 업체의 지역·코스·이용 정보를 한눈에 확인할 수 있는 VIP 추천입니다. 원하는 일정에 맞춰 상세 안내를 확인해보세요.","sort":1},{"shop":"오늘밤테라피","region":"서울 · 경기 · 인천","price":"60,000원~","image":"/assets/images/main-recommend/vip-02.svg","link":"/gangnam-gu/역삼동/shops/001/","desc":"오늘밤테라피 업체의 지역·코스·이용 정보를 한눈에 확인할 수 있는 VIP 추천입니다. 원하는 일정에 맞춰 상세 안내를 확인해보세요.","sort":2},{"shop":"쥬쥬테라피","region":"서울 · 경기 · 인천","price":"60,000원~","image":"/assets/images/main-recommend/vip-03.svg","link":"/gangnam-gu/역삼동/shops/002/","desc":"쥬쥬테라피 업체의 지역·코스·이용 정보를 한눈에 확인할 수 있는 VIP 추천입니다. 원하는 일정에 맞춰 상세 안내를 확인해보세요.","sort":3},{"shop":"퀸즈테라피","region":"서울 · 경기 · 인천","price":"60,000원~","image":"/assets/images/main-recommend/vip-04.svg","link":"/gangnam-gu/역삼동/shops/003/","desc":"퀸즈테라피 업체의 지역·코스·이용 정보를 한눈에 확인할 수 있는 VIP 추천입니다. 원하는 일정에 맞춰 상세 안내를 확인해보세요.","sort":4},{"shop":"한국인골든테라피","region":"서울 · 경기 · 인천","price":"110,000원~","image":"/assets/images/main-recommend/vip-05.png","link":"/gangnam-gu/역삼동/shops/005/","desc":"한국인골든테라피의 지역별 코스와 이용 정보를 확인할 수 있는 VIP 추천 업체입니다.","sort":5}],"premium":[{"shop":"S슬립테라피","region":"대전 · 청주 · 공주","price":"50,000원~","image":"/assets/images/main-recommend/premium-02.svg","link":"/daejeon/daedeok-gu/대화동/shops/001/","desc":"S슬립테라피 업체의 이용 조건과 코스 정보를 편리하게 비교할 수 있는 프리미엄 추천입니다. 상세페이지에서 필요한 내용을 확인해보세요.","sort":2},{"shop":"레이디테라피","region":"대전","price":"60,000원~","image":"/assets/images/main-recommend/premium-03.svg","link":"/daejeon/daedeok-gu/대화동/shops/002/","desc":"레이디테라피 업체의 이용 조건과 코스 정보를 편리하게 비교할 수 있는 프리미엄 추천입니다. 상세페이지에서 필요한 내용을 확인해보세요.","sort":3},{"shop":"로제테라피","region":"청주","price":"60,000원~","image":"/assets/images/main-recommend/premium-04.svg","link":"/cheongju/sangdang-gu/가덕면/shops/002/","desc":"로제테라피 업체의 이용 조건과 코스 정보를 편리하게 비교할 수 있는 프리미엄 추천입니다. 상세페이지에서 필요한 내용을 확인해보세요.","sort":4}]};

(function(){
  const D = window.ML_MAIN_RECOMMEND || {vip:[],premium:[]};
  const esc = s => String(s ?? "")
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");

  function card(x,badge){
    return 
      <a class="mlr-card" href="">
        <div class="mlr-thumb">
          <img src="" alt="">
          <span></span>
          <b></b>
        </div>
        <div class="mlr-body">
          <h3></h3>
          <p></p>
          <div class="mlr-foot">
            <strong></strong>
            <em>상세보기</em>
          </div>
        </div>
      </a>;
  }

  function group(title,badge,items){
    if(!items || !items.length) return "";
    return 
      <section class="mlr-group">
        <h2></h2>
        <div class="mlr-grid"></div>
      </section>;
  }

  function style(){
    if(document.getElementById("mlr-style")) return;
    const s=document.createElement("style");
    s.id="mlr-style";
    s.textContent=
      #ml-main-recommend-final{margin:28px 0 38px}
      #ml-main-recommend-final .mlr-group{margin:0 0 34px}
      #ml-main-recommend-final .mlr-group h2{margin:0 0 14px;font-size:22px;color:#111}
      #ml-main-recommend-final .mlr-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}
      #ml-main-recommend-final .mlr-card{overflow:hidden;border:1px solid #ddd;border-radius:14px;background:#fff;text-decoration:none;color:#111;box-shadow:0 2px 8px rgba(0,0,0,.05)}
      #ml-main-recommend-final .mlr-thumb{height:120px;position:relative;overflow:hidden;background:#174f43}
      #ml-main-recommend-final .mlr-thumb img{width:100%;height:100%;object-fit:cover;display:block}
      #ml-main-recommend-final .mlr-thumb span{position:absolute;left:0;top:0;background:#efc342;color:#111;padding:14px 20px;border-radius:0 0 14px 0;font-size:11px;font-weight:900}
      #ml-main-recommend-final .mlr-thumb b{position:absolute;left:12px;bottom:10px;color:#fff;font-size:12px;text-shadow:0 1px 3px #000}
      #ml-main-recommend-final .mlr-body{padding:12px}
      #ml-main-recommend-final .mlr-body h3{margin:0 0 6px;font-size:15px}
      #ml-main-recommend-final .mlr-body p{height:38px;overflow:hidden;margin:0 0 14px;color:#777;font-size:11px;line-height:1.55}
      #ml-main-recommend-final .mlr-foot{display:flex;align-items:center;justify-content:space-between;gap:8px}
      #ml-main-recommend-final .mlr-foot strong{font-size:14px}
      #ml-main-recommend-final .mlr-foot em{font-style:normal;border:1px solid #e5ad32;border-radius:8px;padding:6px 9px;font-size:10px;font-weight:800}
      @media(max-width:760px){#ml-main-recommend-final .mlr-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
    ;
    document.head.appendChild(s);
  }

  function removePrevious(){
    document.getElementById("ml-main-recommend-auto")?.remove();
    document.getElementById("ml-main-recommend-final")?.remove();
  }

  function render(){
    removePrevious();
    style();

    const root=document.createElement("div");
    root.id="ml-main-recommend-final";
    root.innerHTML=
      group("VIP 추천","VIP",D.vip)+
      group("프리미엄 추천","PICK",D.premium);

    const main=document.querySelector("main") || document.body;
    const partner=main.querySelector(".partner");

    if(partner) main.insertBefore(root,partner);
    else main.appendChild(root);
  }

  if(document.readyState==="loading"){
    document.addEventListener("DOMContentLoaded",render);
  }else{
    render();
  }
})();

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
    /* === JEONGUK-THERAPY MAIN IMAGE LINK START === */
    const vipImages = [
      "/assets/images/main-recommend/vip-01.png",
      "/assets/images/main-recommend/vip-02.png",
      "/assets/images/main-recommend/vip-03.png",
      "/assets/images/main-recommend/vip-04.png"
    ];
    const premiumImages = [
      "/assets/images/main-recommend/premium-01.png",
      "/assets/images/main-recommend/premium-02.png",
      "/assets/images/main-recommend/premium-03.png",
      "/assets/images/main-recommend/premium-04.png"
    ];
    const imageSet = badge === "VIP" ? vipImages : premiumImages;
    const renderItems = (items || []).slice(0,4).map((x,i)=>({
      ...x,
      image: imageSet[i % imageSet.length]
    }));
    /* === JEONGUK-THERAPY MAIN IMAGE LINK END === */

    const sec = document.createElement("section");
    sec.id = id;
    sec.className = "ml-live-section";
    sec.innerHTML = `
      <div class="ml-live-head">
        <h2>${esc(title)}</h2>
        <span>전체보기</span>
      </div>
      <div class="ml-live-grid">${renderItems.map(x=>card(x,badge)).join("")}</div>`;    return sec;
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


