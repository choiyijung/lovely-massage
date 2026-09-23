(function(){
  function detailCount(el){
    if(!el) return 0;
    const t = (el.innerText || "").replace(/\s+/g," ");
    return (t.match(/상세\s*보기/g) || []).length;
  }

  function findHeading(title){
    const all = Array.from(document.querySelectorAll("h1,h2,h3,h4,strong,b,div"));
    return all.find(el=>{
      if(el.closest("#ml-main-recommend-auto")) return false;
      const own = (el.textContent || "").replace(/\s+/g," ").trim();
      return own === title;
    }) || null;
  }

  function findOldBlock(title){
    const heading = findHeading(title);
    if(!heading) return null;

    let node = heading;
    let best = null;

    while(node && node !== document.body && node !== document.documentElement){
      if(node.id === "ml-main-recommend-auto") return null;

      const txt = (node.innerText || "").replace(/\s+/g," ");
      const count = detailCount(node);

      if(count >= 2){
        best = node;
        break;
      }

      // 너무 큰 부모(main 등)까지 올라가지 않도록 제한
      if(node.tagName === "MAIN") break;
      node = node.parentElement;
    }

    return best;
  }

  function apply(){
    const vip = findOldBlock("VIP 추천");
    const premium = findOldBlock("프리미엄 추천");
    const auto = document.getElementById("ml-main-recommend-auto");

    // 새 추천 영역이 있으면 기존 VIP 자리로 이동
    if(auto){
      const target = vip || premium;
      if(target && target.parentNode && auto !== target.previousElementSibling){
        target.parentNode.insertBefore(auto,target);
      }
      auto.style.setProperty("display","block","important");
    }

    [vip,premium].forEach(block=>{
      if(!block) return;
      if(block.id === "ml-main-recommend-auto") return;
      block.style.setProperty("display","none","important");
      block.setAttribute("data-ml-old-recommend","hidden");
    });

    // 임시 업체명이 남은 별도 카드 묶음도 숨김
    const fakeNames = ["강남 골드케어","송파 밸런스룸","광교 리셋테라피","부평 온기케어"];
    document.querySelectorAll("section,div").forEach(el=>{
      if(el.closest("#ml-main-recommend-auto")) return;
      const txt = (el.innerText || "").replace(/\s+/g," ");
      const hits = fakeNames.filter(n=>txt.includes(n)).length;
      if(hits >= 3 && detailCount(el) >= 2){
        el.style.setProperty("display","none","important");
        el.setAttribute("data-ml-old-recommend","hidden");
      }
    });
  }

  function boot(){
    apply();
    requestAnimationFrame(apply);
    setTimeout(apply,100);
    setTimeout(apply,500);
    setTimeout(apply,1500);

    const obs = new MutationObserver(()=>apply());
    obs.observe(document.body,{childList:true,subtree:true});
    setTimeout(()=>obs.disconnect(),5000);
  }

  if(document.readyState === "loading"){
    document.addEventListener("DOMContentLoaded",boot);
  }else{
    boot();
  }
})();