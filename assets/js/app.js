      const vipShops = [{"name":"한국미인테라피","area":"서울 경기 인천","address":"한국미인테라피 지역별 코스와 가격, 이용 정보를 한눈에 확인할 수 있습니다. 상세페이지에서 자세한 이용 안내를 확인해보세요.","price":"100,000원~","image":"assets/images/main-recommend/vip-01.png","link":"/gangnam-gu/shops/001/","type":"VIP","sort":1},{"name":"오늘밤테라피","area":"서울 경기 인천","address":"오늘밤테라피 지역별 코스와 가격, 이용 정보를 한눈에 확인할 수 있습니다. 상세페이지에서 자세한 이용 안내를 확인해보세요.","price":"60,000원~","image":"assets/images/main-recommend/vip-02.png","link":"/gangnam-gu/역삼동/shops/001/","type":"VIP","sort":2},{"name":"쥬쥬테라피","area":"서울 경기 인천","address":"쥬쥬테라피 지역별 코스와 가격, 이용 정보를 한눈에 확인할 수 있습니다. 상세페이지에서 자세한 이용 안내를 확인해보세요.","price":"60,000원~","image":"assets/images/main-recommend/vip-03.png","link":"/gangnam-gu/역삼동/shops/002/","type":"VIP","sort":3},{"name":"퀸즈테라피","area":"서울 경기 인천","address":"퀸즈테라피 지역별 코스와 가격, 이용 정보를 한눈에 확인할 수 있습니다. 상세페이지에서 자세한 이용 안내를 확인해보세요.","price":"60,000원~","image":"assets/images/main-recommend/vip-04.png","link":"/gangnam-gu/역삼동/shops/003/","type":"VIP","sort":4},{"name":"한국인골든테라피","area":"서울 경기 인천","address":"한국인골든테라피의 지역별 코스와 이용 정보를 확인할 수 있는 VIP 추천 업체입니다.","price":"110,000원~","image":"assets/images/main-recommend/vip-05.png","link":"/gangnam-gu/역삼동/shops/005/","type":"VIP","sort":5}];
      const premiumShops = [{"name":"S슬립테라피","area":"대전 청주 공주","address":"S슬립테라피 코스와 가격, 이용 정보를 편리하게 확인할 수 있습니다. 상세페이지에서 자세한 안내를 확인해보세요.","price":"50,000원~","image":"assets/images/main-recommend/premium-02.png","link":"/daejeon/daedeok-gu/대화동/shops/001/","type":"PICK","sort":2},{"name":"레이디테라피","area":"대전","address":"레이디테라피 코스와 가격, 이용 정보를 편리하게 확인할 수 있습니다. 상세페이지에서 자세한 안내를 확인해보세요.","price":"60,000원~","image":"assets/images/main-recommend/premium-03.png","link":"/daejeon/daedeok-gu/대화동/shops/002/","type":"PICK","sort":3},{"name":"로제테라피","area":"청주","address":"로제테라피 코스와 가격, 이용 정보를 편리하게 확인할 수 있습니다. 상세페이지에서 자세한 안내를 확인해보세요.","price":"60,000원~","image":"assets/images/main-recommend/premium-04.png","link":"/cheongju/sangdang-gu/가덕면/shops/002/","type":"PICK","sort":4}];

      const shops = [...vipShops, ...premiumShops];

const card = (s, type, idx=0) => `
<article class="card">
  <div class="thumb">
    <img src="${s.image || ('assets/images/shops/room-' + ((idx % 8) + 1) + '.jpg')}" alt="${s.name} 추천 이미지">
    <span class="badge" data-label="${type}" aria-label="${type}"></span>
    <div class="thumb-title">${s.area}</div>
  </div>
  <div class="card-body">
    <h3>${s.name}</h3>
    <div class="card-sub">${s.area}</div>
    <div class="stats"><span>${type}</span><span>테라피 안내</span></div>
    
    <div class="address">${s.address}</div>
    <div class="card-bottom">
      <div class="price">${s.price}</div>
            <a href="${s.link || '#'}" class="detail-btn">자세히 보기</a>
    </div>
  </div>
</article>`;

function mlShufflePerDevice(list,key){
  const base=[...list];
  const prevKey=key+"-prev";
  const deviceKey=key+"-device";

  let deviceId=localStorage.getItem(deviceKey);
  if(!deviceId){
    const id=new Uint32Array(4);
    crypto.getRandomValues(id);
    deviceId=Array.from(id,n=>n.toString(16).padStart(8,"0")).join("");
    localStorage.setItem(deviceKey,deviceId);
  }

  let deviceHash=2166136261;
  for(const c of deviceId){
    deviceHash^=c.charCodeAt(0);
    deviceHash=Math.imul(deviceHash,16777619)>>>0;
  }

  const randomInt=(max)=>{
    const a=new Uint32Array(1);
    crypto.getRandomValues(a);
    return ((a[0]^deviceHash)>>>0)%max;
  };

  const prev=localStorage.getItem(prevKey)||"";
  let out=[...base];
  let sig="";

  for(let attempt=0;attempt<30;attempt++){
    out=[...base];

    for(let i=out.length-1;i>0;i--){
      const j=randomInt(i+1);
      [out[i],out[j]]=[out[j],out[i]];
    }

    sig=out.map(x=>x.name||x.shop||"").join("|");

    if(sig!==prev) break;
  }

  if(sig===prev && out.length>1){
    out.push(out.shift());
    sig=out.map(x=>x.name||x.shop||"").join("|");
  }

  localStorage.setItem(prevKey,sig);
  return out;
}

document.querySelector("#vipCards").innerHTML =
  mlShufflePerDevice(vipShops,"ml-vip-order")
    .map((s,i)=>card(s,"VIP",i))
    .join("");

document.querySelector("#premiumCards").innerHTML = premiumShops.map((s,i)=>card(s,"PICK",i+4)).join("");

const regions = {"서울":[{"name":"한국미인테라피","area":"서울","price":"100,000원~","link":"/gangnam-gu/shops/001/"},{"name":"오늘밤테라피","area":"서울","price":"60,000원~","link":"/gangnam-gu/역삼동/shops/001/"},{"name":"쥬쥬테라피","area":"서울","price":"60,000원~","link":"/gangnam-gu/역삼동/shops/002/"},{"name":"한국인골든테라피","area":"서울","price":"110,000원~","link":"/gangnam-gu/역삼동/shops/005/"},{"name":"퀸즈테라피","area":"서울","price":"60,000원~","link":"/gangnam-gu/역삼동/shops/003/"}],"경기":[{"name":"한국미인테라피","area":"경기","price":"100,000원~","link":"/ansan/danwon-gu/고잔동/shops/004/"},{"name":"오늘밤테라피","area":"경기","price":"60,000원~","link":"/ansan/danwon-gu/고잔동/shops/001/"},{"name":"쥬쥬테라피","area":"경기","price":"60,000원~","link":"/ansan/danwon-gu/고잔동/shops/002/"},{"name":"한국인골든테라피","area":"경기","price":"110,000원~","link":"/ansan/danwon-gu/고잔동/shops/005/"},{"name":"퀸즈테라피","area":"경기","price":"60,000원~","link":"/ansan/danwon-gu/고잔동/shops/003/"}],"인천":[{"name":"한국미인테라피","area":"인천","price":"100,000원~","link":"/ganghwa-gun/강화읍/shops/004/"},{"name":"오늘밤테라피","area":"인천","price":"60,000원~","link":"/ganghwa-gun/강화읍/shops/001/"},{"name":"쥬쥬테라피","area":"인천","price":"60,000원~","link":"/ganghwa-gun/강화읍/shops/002/"},{"name":"한국인골든테라피","area":"인천","price":"110,000원~","link":"/ganghwa-gun/강화읍/shops/005/"},{"name":"퀸즈테라피","area":"인천","price":"60,000원~","link":"/ganghwa-gun/강화읍/shops/003/"}],"기타지역":[{"name":"한국미인테라피","area":"천안 아산","price":"100,000원~","link":"/asan/도고면/shops/002/"},{"name":"쥬쥬테라피","area":"천안 아산","price":"60,000원~","link":"/asan/도고면/shops/001/"},{"name":"S슬립테라피","area":"대전 청주 공주","price":"50,000원~","link":"/daejeon/daedeok-gu/대화동/shops/001/"},{"name":"레이디테라피","area":"대전","price":"60,000원~","link":"/daejeon/daedeok-gu/대화동/shops/002/"}]};

function drawRegion(key){
  /* === CLEAN-MASSAGE REGION IMAGE START === */
  const regionImages = [
    "assets/images/main-recommend/region-01.png",
    "assets/images/main-recommend/region-02.png",
    "assets/images/main-recommend/region-03.png",
    "assets/images/main-recommend/region-04.png",
    "assets/images/main-recommend/region-05.png",
    "assets/images/main-recommend/region-06.png",
    "assets/images/main-recommend/region-07.png",
    "assets/images/main-recommend/region-08.png"
  ];
  const tabOffset = { "서울":0, "경기":2, "인천":4, "기타지역":6 };
  const offset = tabOffset[key] || 0;

  document.querySelector("#regionItems").innerHTML = regions[key].map((x,i)=>`
    <a class="region-row region-row-with-image" href="${x.link || '#'}" style="text-decoration:none;color:inherit">
      <span class="region-row-thumb">
        <img src="${regionImages[(i + offset) % regionImages.length]}" alt="${x.area} ${x.name}" loading="lazy">
        <b>${i+1}</b>
      </span>
      <div>
        <h3>${x.name}</h3>
        <p>${x.area} 등록 업체</p>
      </div>
      <strong>${x.price}</strong>
    </a>`).join("");
  /* === CLEAN-MASSAGE REGION IMAGE END === */
}
drawRegion("서울");

document.querySelectorAll(".tabs button").forEach(btn=>{
  btn.addEventListener("click",()=>{
    document.querySelectorAll(".tabs button").forEach(x=>x.classList.remove("active"));
    btn.classList.add("active");
    drawRegion(btn.dataset.tab);
  });
});

const drawer = document.querySelector("#drawer");
const overlay = document.querySelector("#overlay");

function menu(open){
  drawer.classList.toggle("open",open);
  overlay.classList.toggle("show",open);
  drawer.setAttribute("aria-hidden",String(!open));
}
document.querySelector("#openMenu").addEventListener("click",()=>menu(true));
document.querySelector("#closeMenu").addEventListener("click",()=>menu(false));
overlay.addEventListener("click",()=>menu(false));

document.querySelector("#searchForm").addEventListener("submit",(e)=>{
  e.preventDefault();
  const q = document.querySelector("#q").value.trim().toLowerCase();
  if(!q) return;
  document.querySelectorAll(".card,.region-row").forEach(el=>{
    const hit = el.textContent.toLowerCase().includes(q);
    el.style.outline = hit ? "2px solid #d8aa4a" : "";
    el.style.outlineOffset = hit ? "2px" : "";
  });
  document.querySelector("#vipCards").scrollIntoView({behavior:"smooth",block:"center"});
});

/* ===== v26 badge dedupe ===== */
function normalizeCardBadges(){
  document.querySelectorAll(".thumb").forEach((thumb)=>{
    const badges = Array.from(thumb.children).filter((el)=>el.classList && el.classList.contains("badge"));
    badges.forEach((badge,index)=>{
      if(index>0) badge.remove();
    });
  });
}
normalizeCardBadges();
requestAnimationFrame(normalizeCardBadges);



