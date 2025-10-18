
// 点击展开/收起子菜单
document.querySelectorAll(".has-submenu").forEach(item => {
  item.addEventListener("click", function (e) {
    e.stopPropagation();
    this.classList.toggle("active");
  });
});

// 点击具体“表”时，加载对应iframe + 显示分析面板
document.querySelectorAll("li[data-url]").forEach(item => {
  item.addEventListener("click", function (e) {
    e.stopPropagation();

    const iframe = document.getElementById("content-frame");
    const placeholder = document.getElementById("placeholder");
    const allPanels = document.querySelectorAll(".analysis-panel");

    // 隐藏所有分析区
    allPanels.forEach(p => p.style.display = "none");

    // 设置iframe内容
    const url = this.getAttribute("data-url");
    iframe.src = url;
    iframe.style.display = "block";
    placeholder.style.display = "none";

    // 显示对应分析区
    const panelId = this.getAttribute("data-panel");
    const panel = document.getElementById(panelId);
    if (panel) panel.style.display = "block";
  });
});

