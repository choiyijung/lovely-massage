window.ML_MAIN_RECOMMEND = {"vip":[{"kind":"vip","shop":"한국미인테라피","region":"서울","price":"100,000원","image":"/assets/images/main-recommend/vip-01.svg","link":"/gangnam-gu/shops/001/","desc":"지역과 이용 조건을 한눈에 확인하기 좋은 추천 업체입니다. 원하는 일정과 코스를 비교한 뒤 편하게 문의해보세요.","sort":1},{"kind":"vip","shop":"오늘밤테라피","region":"강남","price":"60,000원","image":"/assets/images/main-recommend/vip-02.svg","link":"/gangnam-gu/역삼동/shops/001/","desc":"깔끔한 이용 안내와 편리한 예약 확인을 중심으로 살펴볼 수 있는 추천 업체입니다. 위치와 시간을 확인해 선택해보세요.","sort":2},{"kind":"vip","shop":"쥬쥬테라피","region":"인천","price":"60,000원","image":"/assets/images/main-recommend/vip-03.svg","link":"/ganghwa-gun/강화읍/shops/002/","desc":"지역별 이용 범위와 코스 정보를 간편하게 비교할 수 있는 추천 업체입니다. 방문 전 필요한 내용을 확인해보세요.","sort":3},{"kind":"vip","shop":"퀸즈테라피","region":"경기","price":"60,000원","image":"/assets/images/main-recommend/vip-04.svg","link":"/mapo-gu/성산동/shops/003/","desc":"일정에 맞춰 업체 정보와 이용 조건을 빠르게 확인할 수 있는 추천 업체입니다. 상세페이지에서 코스와 안내를 확인해보세요.","sort":4},{"shop":"한국인골든테라피","region":"서울 · 경기 · 인천","price":"110,000원~","image":"/assets/images/main-recommend/vip-05.png","link":"/gangnam-gu/역삼동/shops/005/","desc":"한국인골든테라피의 지역별 코스와 이용 정보를 확인할 수 있는 VIP 추천 업체입니다.","sort":5}],"premium":[{"kind":"premium","shop":"S슬립테라피","region":"대전","price":"50,000원","image":"/assets/images/main-recommend/premium-02.svg","link":"/daejeon/daedeok-gu/대화동/shops/001/","desc":"업체 정보와 이용 조건을 보기 쉽게 확인할 수 있는 프리미엄 추천 업체입니다. 원하는 일정에 맞춰 살펴보세요.","sort":2},{"kind":"premium","shop":"레이디테라피","region":"대전","price":"60,000원","image":"/assets/images/main-recommend/premium-03.svg","link":"/daejeon/daedeok-gu/대화동/shops/002/","desc":"지역, 가격, 코스 정보를 한 번에 비교하기 좋은 프리미엄 추천 업체입니다. 상세 안내를 확인한 뒤 선택해보세요.","sort":3},{"kind":"premium","shop":"상큼테라피","region":"전주","price":"60,000원","image":"/assets/images/main-recommend/premium-04.svg","link":"/jeonju/deokjin-gu/금암동/shops/002/","desc":"편리한 예약 확인과 명확한 이용 안내를 기준으로 살펴볼 수 있는 프리미엄 추천 업체입니다. 필요한 정보를 확인해보세요.","sort":4}]};

(function(){
  const DATA = window.ML_MAIN_RECOMMEND || {vip:[], premium:[]};

  function esc(v){
    return String(v == null ? "" : v)
      .replace(/&/g,"&amp;")
      .replace(/</g,"&lt;")
      .replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;");
  }

  function money(v){
    const s = String(v || "").trim();
    if(!s) return "가격 안내";
    if(/^\d+$/.test(s)) return Number(s).toLocaleString("ko-KR") + "원";
    return s;
  }

  function hideOldRecommend(){
    const fakeNames = ["강남 골드케어","송파 밸런스룸","광교 리셋테라피","부평 온기케어"];
    document.querySelectorAll("section").forEach(sec=>{
      if(sec.id === "ml-main-recommend-auto") return;
      const t = (sec.innerText || "").replace(/\s+/g," ");
      const oldVip = /VIP\s*추천/i.test(t);
      const oldPremium = /프리미엄\s*추천/.test(t);
      const fake = fakeNames.some(n=>t.includes(n));
      if(oldVip || oldPremium || fake){
        sec.style.setProperty("display","none","important");
        sec.setAttribute("data-ml-old-recommend","hidden");
      }
    });
  }

  function card(item, label){
    const link = item.link && item.link !== "#" ? item.link : "javascript:void(0)";
    const disabled = item.link && item.link !== "#" ? "" : ' aria-disabled="true"';
    return 
      <a class="ml-rec-card" href="">
        <div class="ml-rec-image">
          <img src="" alt=" 추천 이미지" loading="lazy">
          <span></span>
        </div>
        <div class="ml-rec-body">
          <div class="ml-rec-region"></div>
          <h3></h3>
          <p></p>
          <div class="ml-rec-bottom">
            <b></b>
            <em>상세보기</em>
          </div>
        </div>
      </a>;
  }

  function block(title, sub, items, label){
    if(!items || !items.length) return "";
    return 
      <section class="ml-rec-group">
        <div class="ml-rec-heading">
          <div>
            <small></small>
            <h2></h2>
            <p></p>
          </div>
        </div>
        <div class="ml-rec-grid">
          
        </div>
      </section>;
  }

  function ensureStyle(){
    if(document.getElementById("ml-main-recommend-style")) return;
    const style = document.createElement("style");
    style.id = "ml-main-recommend-style";
    style.textContent = 
      #ml-main-recommend-auto{margin:34px 0 42px}
      #ml-main-recommend-auto .ml-rec-group{margin:0 0 34px}
      #ml-main-recommend-auto .ml-rec-heading{display:flex;align-items:end;justify-content:space-between;margin:0 0 16px}
      #ml-main-recommend-auto .ml-rec-heading small{display:inline-block;color:#aa7a00;font-size:12px;font-weight:900;letter-spacing:.08em;margin-bottom:5px}
      #ml-main-recommend-auto .ml-rec-heading h2{margin:0;color:#163b32;font-size:27px;line-height:1.2}
      #ml-main-recommend-auto .ml-rec-heading p{margin:7px 0 0;color:#75827e;font-size:14px}
      #ml-main-recommend-auto .ml-rec-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px}
      #ml-main-recommend-auto .ml-rec-card{display:grid;grid-template-columns:44% 56%;min-height:190px;overflow:hidden;border:1px solid #e5e9e7;border-radius:18px;background:#fff;color:inherit;text-decoration:none;box-shadow:0 8px 26px rgba(15,53,44,.07);transition:.18s ease}
      #ml-main-recommend-auto .ml-rec-card:hover{transform:translateY(-2px);box-shadow:0 12px 30px rgba(15,53,44,.12)}
      #ml-main-recommend-auto .ml-rec-image{position:relative;min-height:190px;background:#0b4a3e;overflow:hidden}
      #ml-main-recommend-auto .ml-rec-image img{width:100%;height:100%;object-fit:cover;display:block}
      #ml-main-recommend-auto .ml-rec-image span{position:absolute;left:12px;top:12px;background:#ffc928;color:#083f35;border-radius:999px;padding:6px 10px;font-size:11px;font-weight:900}
      #ml-main-recommend-auto .ml-rec-body{padding:20px 20px 16px;display:flex;flex-direction:column}
      #ml-main-recommend-auto .ml-rec-region{font-size:12px;font-weight:800;color:#a57800;margin-bottom:5px}
      #ml-main-recommend-auto h3{margin:0 0 9px;color:#163b32;font-size:20px}
      #ml-main-recommend-auto .ml-rec-body p{margin:0;color:#6f7c78;font-size:13px;line-height:1.55;flex:1}
      #ml-main-recommend-auto .ml-rec-bottom{display:flex;align-items:center;justify-content:space-between;margin-top:14px;gap:10px}
      #ml-main-recommend-auto .ml-rec-bottom b{color:#163b32;font-size:15px}
      #ml-main-recommend-auto .ml-rec-bottom em{font-style:normal;background:#ffc928;color:#083f35;border-radius:8px;padding:7px 11px;font-size:11px;font-weight:900}
      @media(max-width:760px){
        #ml-main-recommend-auto{margin:24px 0 32px}
        #ml-main-recommend-auto .ml-rec-grid{grid-template-columns:1fr}
        #ml-main-recommend-auto .ml-rec-card{grid-template-columns:39% 61%;min-height:165px}
        #ml-main-recommend-auto .ml-rec-image{min-height:165px}
        #ml-main-recommend-auto .ml-rec-body{padding:15px}
        #ml-main-recommend-auto h3{font-size:18px}
        #ml-main-recommend-auto .ml-rec-heading h2{font-size:23px}
      };
    document.head.appendChild(style);
  }

  function render(){
    hideOldRecommend();
    ensureStyle();

    let root = document.getElementById("ml-main-recommend-auto");
    if(!root){
      root = document.createElement("div");
      root.id = "ml-main-recommend-auto";

      const main = document.querySelector("main") || document.body;
      const partner = main.querySelector(".partner");
      const guide = Array.from(main.querySelectorAll("section")).find(s=>{
        const t = (s.innerText || "");
        return /이용안내|지역정보/.test(t);
      });

      const anchor = guide || partner;
      if(anchor && anchor.parentNode === main) main.insertBefore(root, anchor);
      else main.appendChild(root);
    }

    root.innerHTML =
      block("VIP 추천","메인에서 별도로 관리하는 VIP 추천 업체입니다.",DATA.vip,"VIP 추천") +
      block("프리미엄 추천","메인에서 별도로 관리하는 프리미엄 추천 업체입니다.",DATA.premium,"PREMIUM");
  }

  if(document.readyState === "loading"){
    document.addEventListener("DOMContentLoaded",()=>{
      render();
      requestAnimationFrame(render);
      setTimeout(render,120);
    });
  }else{
    render();
    requestAnimationFrame(render);
    setTimeout(render,120);
  }
})();