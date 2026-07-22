/**
 * 伪装模式 — 党史学习资料 × Microsoft Word
 *
 * 快捷键：
 *   Ctrl+Shift+D  — 开关伪装模式
 *   `（反引号）   — 紧急掩护（仅替换侧栏+正文，再次关闭）
 *   连按两次 Esc  — 紧急掩护
 *
 * 状态持久化：localStorage['dol_disguise_mode']
 * URL 参数：?disguise=1 启动即开；?disguise=0 强制关
 */
(function () {
  "use strict";

  var STORAGE_KEY = "dol_disguise_mode";
  var TITLE_DISGUISE = "新时代中国特色社会主义思想学习纲要（修订本）.docx - Word";
  var TITLE_FALLBACK = "Degrees of Lewdity";
  var originalTitle = document.title || TITLE_FALLBACK;
  var panicVisible = false;
  var lastEscAt = 0;
  /** #stats 移到底栏前的原位置，用于退出伪装时还原 */
  var statsHomeParent = null;
  var statsHomeNext = null;

  function isOn() {
    return document.documentElement.classList.contains("disguise-mode");
  }

  function setOn(on) {
    on = !!on;
    document.documentElement.classList.toggle("disguise-mode", on);
    if (document.body) {
      document.body.classList.toggle("disguise-mode", on);
    }

    try {
      localStorage.setItem(STORAGE_KEY, on ? "1" : "0");
    } catch (e) { /* ignore */ }

    if (on) {
      originalTitle = document.title && document.title !== TITLE_DISGUISE
        ? document.title
        : originalTitle;
      document.title = TITLE_DISGUISE;
      ensureChrome();
      paintChrome();
      relocateStatsToStatusbar();
      fixNextButtons();
      updateStatusbar();
      scrubLightTextColors(document.getElementById("passages"));
      scrubLightTextColors(document.getElementById("ui-bar"));
      scrubLightTextColors(document.getElementById("story-caption"));
      scrubLightTextColors(document.getElementById("storyCaptionDiv"));
      scrubLightTextColors(document.getElementById("storyCaptionContent"));
    } else {
      restoreStatsToSidebar();
      clearNextButtonFixes();
      document.title = originalTitle || TITLE_FALLBACK;
      hidePanic();
    }

    var fab = document.getElementById("disguise-fab");
    if (fab) {
      fab.setAttribute("aria-pressed", on ? "true" : "false");
      fab.title = on
        ? "退出伪装模式 (Ctrl+Shift+D)"
        : "伪装模式：党史×Word (Ctrl+Shift+D)";
    }
  }

  function toggle() {
    setOn(!isOn());
  }

  function paintChrome() {
    var title = document.getElementById("disguise-titlebar");
    if (title) {
      title.style.setProperty("background", "#2b579a", "important");
      title.style.setProperty("background-color", "#2b579a", "important");
      title.style.setProperty("background-image", "none", "important");
      title.style.setProperty("color", "#ffffff", "important");
    }
    var tabs = document.getElementById("disguise-ribbon-tabs");
    if (tabs) {
      tabs.style.setProperty("background", "#f3f2f1", "important");
      tabs.style.setProperty("background-color", "#f3f2f1", "important");
      tabs.style.setProperty("background-image", "none", "important");
      tabs.style.setProperty("color", "#252525", "important");
      var tabNodes = tabs.querySelectorAll(".tab");
      for (var i = 0; i < tabNodes.length; i++) {
        var t = tabNodes[i];
        var active = t.classList.contains("active");
        t.style.setProperty("background", active ? "#ffffff" : "transparent", "important");
        t.style.setProperty("background-color", active ? "#ffffff" : "transparent", "important");
        t.style.setProperty("color", active ? "#2b579a" : "#252525", "important");
        t.style.setProperty("-webkit-text-fill-color", active ? "#2b579a" : "#252525", "important");
      }
    }
    var ribbon = document.getElementById("disguise-ribbon-body");
    if (ribbon) {
      ribbon.style.setProperty("background", "#ffffff", "important");
      ribbon.style.setProperty("background-color", "#ffffff", "important");
      ribbon.style.setProperty("color", "#252525", "important");
    }
    var status = document.getElementById("disguise-statusbar");
    if (status) {
      status.style.setProperty("background", "#2b579a", "important");
      status.style.setProperty("background-color", "#2b579a", "important");
      status.style.setProperty("color", "#ffffff", "important");
    }
    // 清掉可能残留的旧 banner
    var oldBanner = document.getElementById("disguise-banner");
    if (oldBanner && oldBanner.parentNode) oldBanner.parentNode.removeChild(oldBanner);
  }

  function ensureChrome() {
    if (document.getElementById("disguise-chrome")) {
      // 旧版底栏无 stats 槽时补建
      var status = document.getElementById("disguise-statusbar");
      if (status && !document.getElementById("disguise-stats-slot")) {
        var slot = document.createElement("div");
        slot.id = "disguise-stats-slot";
        slot.title = "点击展开/收起状态";
        status.insertBefore(slot, status.firstChild);
      }
      paintChrome();
      return;
    }
    if (!document.body) return;

    var chrome = document.createElement("div");
    chrome.id = "disguise-chrome";
    chrome.setAttribute("aria-hidden", "true");
    chrome.innerHTML =
      '<div id="disguise-titlebar">' +
        '<div class="doc-icon"></div>' +
        '<div class="doc-title">新时代中国特色社会主义思想学习纲要（修订本）.docx - Word</div>' +
        '<div class="win-btns"><span>—</span><span>□</span><span>×</span></div>' +
      "</div>" +
      '<div id="disguise-ribbon-tabs">' +
        '<div class="tab">文件</div>' +
        '<div class="tab active">开始</div>' +
        '<div class="tab">插入</div>' +
        '<div class="tab">设计</div>' +
        '<div class="tab">布局</div>' +
        '<div class="tab">引用</div>' +
        '<div class="tab">审阅</div>' +
        '<div class="tab">视图</div>' +
      "</div>" +
      '<div id="disguise-ribbon-body">' +
        '<div class="group">' +
          '<div class="group-tools">' +
            '<div class="btn"><div class="ico">📋</div>粘贴</div>' +
            '<div class="btn-row">' +
              '<div class="btn-sm">✂ 剪切</div>' +
              '<div class="btn-sm">📄 复制</div>' +
              '<div class="btn-sm">🖌 格式刷</div>' +
            "</div>" +
          "</div>" +
          '<div class="group-label">剪贴板</div>' +
        "</div>" +
        '<div class="group">' +
          '<div class="group-tools">' +
            '<div class="font-box">' +
              '<div class="font-row">' +
                '<div class="fake-select">宋体</div>' +
                '<div class="fake-select size">12</div>' +
              "</div>" +
              '<div class="font-row">' +
                '<div class="fmt b">B</div>' +
                '<div class="fmt i">I</div>' +
                '<div class="fmt u">U</div>' +
                '<div class="fmt red">A</div>' +
                '<div class="fmt">≡</div>' +
              "</div>" +
            "</div>" +
          "</div>" +
          '<div class="group-label">字体</div>' +
        "</div>" +
        '<div class="group">' +
          '<div class="group-tools">' +
            '<div class="btn"><div class="ico">¶</div>段落</div>' +
            '<div class="btn-row">' +
              '<div class="btn-sm">• 项目符号</div>' +
              '<div class="btn-sm">1. 编号</div>' +
              '<div class="btn-sm">⇔ 两端对齐</div>' +
            "</div>" +
          "</div>" +
          '<div class="group-label">段落</div>' +
        "</div>" +
        '<div class="group">' +
          '<div class="group-tools">' +
            '<div class="btn"><div class="ico">📑</div>样式</div>' +
            '<div class="btn-row">' +
              '<div class="btn-sm heading1">标题 1</div>' +
              '<div class="btn-sm">标题 2</div>' +
              '<div class="btn-sm">正文</div>' +
            "</div>" +
          "</div>" +
          '<div class="group-label">样式</div>' +
        "</div>" +
        '<div class="group">' +
          '<div class="group-tools">' +
            '<div class="btn"><div class="ico">🔍</div>查找</div>' +
            '<div class="btn"><div class="ico">✏</div>替换</div>' +
            '<div class="btn" id="disguise-ribbon-toggle" title="退出伪装 (Ctrl+Shift+D)"><div class="ico">👁</div>审阅</div>' +
          "</div>" +
          '<div class="group-label">编辑</div>' +
        "</div>" +
      "</div>";

    var status = document.createElement("div");
    status.id = "disguise-statusbar";
    status.innerHTML =
      '<div id="disguise-stats-slot" title="点击展开/收起状态"></div>' +
      '<div class="right">' +
        '<span id="disguise-status-page">第 1 页，共 48 页</span>' +
        '<span class="sep"></span>' +
        '<span id="disguise-exit" title="退出伪装模式 (Ctrl+Shift+D)">校对：完毕</span>' +
      "</div>";

    document.body.insertBefore(chrome, document.body.firstChild);
    document.body.appendChild(status);
    // 紧急掩护层：分别盖在侧栏 / 正文上，不遮挡 Word 外壳
    ensurePanicLayers();

    var exitBtn = document.getElementById("disguise-exit");
    if (exitBtn) {
      exitBtn.addEventListener("click", function (e) {
        e.preventDefault();
        e.stopPropagation();
        setOn(false);
      });
    }

    var ribbonToggle = document.getElementById("disguise-ribbon-toggle");
    if (ribbonToggle) {
      ribbonToggle.addEventListener("click", function (e) {
        e.preventDefault();
        e.stopPropagation();
        setOn(false);
      });
    }

    // 选项卡切换仅视觉反馈
    chrome.querySelectorAll("#disguise-ribbon-tabs .tab").forEach(function (tab) {
      tab.addEventListener("click", function () {
        chrome.querySelectorAll("#disguise-ribbon-tabs .tab").forEach(function (t) {
          t.classList.remove("active");
        });
        tab.classList.add("active");
        paintChrome();
      });
    });

    paintChrome();
  }

  function ensureFab() {
    if (document.getElementById("disguise-fab") || !document.body) return;
    var fab = document.createElement("button");
    fab.id = "disguise-fab";
    fab.type = "button";
    // Word 文档风格图标（纯 SVG，不依赖外链）
    fab.innerHTML =
      '<svg class="disguise-fab-icon" viewBox="0 0 24 24" width="22" height="22" aria-hidden="true" focusable="false">' +
        '<path fill="#2b579a" d="M6 2h9l5 5v15a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V3a1 1 0 0 1 1-1z"/>' +
        '<path fill="#5b8fd4" d="M15 2v5h5"/>' +
        '<path fill="#fff" d="M8.2 17.2l1.5-6.4h1.55l.95 3.85.95-3.85H14.7l1.5 6.4h-1.45l-.85-3.7-.95 3.7h-1.35l-.95-3.7-.85 3.7H8.2z"/>' +
      "</svg>";
    fab.title = "伪装模式：党史×Word (Ctrl+Shift+D)";
    fab.setAttribute("aria-label", "切换伪装模式");
    fab.setAttribute("aria-pressed", "false");
    fab.addEventListener("click", function (e) {
      e.preventDefault();
      e.stopPropagation();
      toggle();
    });
    document.body.appendChild(fab);
  }

  function updateStatusbar() {
    var el = document.getElementById("disguise-status-page");
    if (!el) return;
    // 用 passage 文本长度生成一个稳定的“页码”观感
    var passages = document.getElementById("passages");
    var len = passages ? (passages.textContent || "").length : 0;
    var page = Math.max(1, Math.min(48, Math.floor(len / 280) + 1));
    el.textContent = "第 " + page + " 页，共 48 页";
    // 段落刷新后 #stats 可能被重建在侧栏，重新挂到底栏
    if (isOn()) {
      relocateStatsToStatusbar();
      fixNextButtons();
    }
  }

  /**
   * 纠正「继续」类固定条（#next/#skip/#stop）：
   * 游戏默认 position:fixed; bottom:-1px，伪装布局下会飞出可视区。
   */
  function fixNextButtons() {
    if (!isOn()) return;
    var ids = ["next", "skip", "stop", "noStop"];
    for (var i = 0; i < ids.length; i++) {
      var el = document.getElementById(ids[i]);
      if (!el) continue;
      el.style.setProperty("position", "sticky", "important");
      el.style.setProperty("bottom", "0", "important");
      el.style.setProperty("left", "auto", "important");
      el.style.setProperty("right", "auto", "important");
      el.style.setProperty("top", "auto", "important");
      el.style.setProperty("transform", "none", "important");
      el.style.setProperty("width", "100%", "important");
      el.style.setProperty("max-width", "100%", "important");
      el.style.setProperty("margin", "1.2em 0 0", "important");
      el.style.setProperty("z-index", "30", "important");
    }
  }

  function clearNextButtonFixes() {
    var ids = ["next", "skip", "stop", "noStop"];
    var props = [
      "position", "bottom", "left", "right", "top", "transform",
      "width", "max-width", "margin", "z-index"
    ];
    for (var i = 0; i < ids.length; i++) {
      var el = document.getElementById(ids[i]);
      if (!el || !el.style) continue;
      for (var j = 0; j < props.length; j++) {
        el.style.removeProperty(props[j]);
      }
    }
  }

  /**
   * 收集文档中所有 id=stats 的节点（游戏重建会产生重复 id）。
   */
  function queryAllStats() {
    var list = [];
    try {
      var nodes = document.querySelectorAll("#stats, [id='stats']");
      for (var i = 0; i < nodes.length; i++) list.push(nodes[i]);
    } catch (e) {
      var one = document.getElementById("stats");
      if (one) list.push(one);
    }
    return list;
  }

  /**
   * 将 storyCaptionDiv 内的 #stats 挂到伪装底栏（始终只保留一份）。
   * 游戏每次刷新 caption 可能在侧栏再插一个 #stats，若不清理会堆在底栏。
   */
  function relocateStatsToStatusbar() {
    var slot = document.getElementById("disguise-stats-slot");
    if (!slot) return;

    var all = queryAllStats();
    if (!all.length) return;

    // 优先使用「不在 slot 内」的最新节点（游戏刚重建的），否则用 slot 内已有的
    var keep = null;
    var i;
    for (i = 0; i < all.length; i++) {
      if (all[i].parentNode !== slot) {
        keep = all[i];
      }
    }
    if (!keep) {
      // 全部已在 slot 或孤立：取最后一个
      keep = all[all.length - 1];
    }

    // 记录还原位置（仅当 keep 仍在侧栏树中时）
    if (keep.parentNode && keep.parentNode !== slot) {
      if (!statsHomeParent || !document.contains(statsHomeParent)) {
        statsHomeParent =
          document.getElementById("storyCaptionDiv") || keep.parentNode;
        statsHomeNext = keep.nextSibling;
      }
    }

    // 删掉其它重复 #stats
    for (i = 0; i < all.length; i++) {
      if (all[i] !== keep && all[i].parentNode) {
        all[i].parentNode.removeChild(all[i]);
      }
    }

    // 清空 slot 里残留的旧节点（防止非 #stats 垃圾或漏网重复）
    while (slot.firstChild) {
      if (slot.firstChild === keep) break;
      slot.removeChild(slot.firstChild);
    }
    // 若 slot 内还有除 keep 以外的子节点，清掉
    var kids = slot.childNodes;
    for (i = kids.length - 1; i >= 0; i--) {
      if (kids[i] !== keep) slot.removeChild(kids[i]);
    }

    if (keep.parentNode !== slot) {
      slot.appendChild(keep);
    }

    // 侧栏不应再残留 #stats
    stripSidebarStats();
    // 底栏不需要模拟钟
    hideClockInStats(keep);
  }

  /** 删除侧栏内残留的 #stats（伪装时只保留底栏一份） */
  function stripSidebarStats() {
    var slot = document.getElementById("disguise-stats-slot");
    var all = queryAllStats();
    for (var i = 0; i < all.length; i++) {
      var el = all[i];
      if (slot && el.parentNode === slot) continue;
      // 在 ui-bar / storyCaption 下的一律移除
      if (
        el.closest &&
        (el.closest("#ui-bar") ||
          el.closest("#story-caption") ||
          el.closest("#storyCaptionDiv"))
      ) {
        if (el.parentNode) el.parentNode.removeChild(el);
      }
    }
  }

  function hideClockInStats(root) {
    if (!root || !root.querySelectorAll) return;
    var clocks = root.querySelectorAll(
      ".clockContainer, .clockCenter, .clockSpinner, .clock"
    );
    for (var i = 0; i < clocks.length; i++) {
      clocks[i].style.setProperty("display", "none", "important");
    }
  }

  function restoreStatsToSidebar() {
    var slot = document.getElementById("disguise-stats-slot");
    var all = queryAllStats();
    var stats = null;

    // 优先取 slot 内那份
    for (var i = 0; i < all.length; i++) {
      if (slot && all[i].parentNode === slot) {
        stats = all[i];
        break;
      }
    }
    if (!stats && all.length) stats = all[0];

    // 去掉多余重复
    for (i = 0; i < all.length; i++) {
      if (all[i] !== stats && all[i].parentNode) {
        all[i].parentNode.removeChild(all[i]);
      }
    }

    if (!stats) {
      statsHomeParent = null;
      statsHomeNext = null;
      return;
    }

    var home = statsHomeParent && document.contains(statsHomeParent)
      ? statsHomeParent
      : document.getElementById("storyCaptionDiv");

    if (!home) {
      statsHomeParent = null;
      statsHomeNext = null;
      return;
    }

    if (
      statsHomeNext &&
      statsHomeNext.parentNode === home &&
      document.contains(statsHomeNext)
    ) {
      home.insertBefore(stats, statsHomeNext);
    } else {
      var content = document.getElementById("storyCaptionContent");
      if (content && content.parentNode === home) {
        home.insertBefore(stats, content);
      } else {
        home.insertBefore(stats, home.firstChild);
      }
    }

    statsHomeParent = null;
    statsHomeNext = null;
  }

  /**
   * 清除过浅的行内颜色（白/近白/亮灰），避免白底看不清。
   * 不碰链接本身，由 CSS 统一成 Word 蓝。
   */
  function isTooLight(colorStr) {
    if (!colorStr) return false;
    var s = String(colorStr).toLowerCase().trim();
    if (s === "white" || s === "#fff" || s === "#ffffff" || s === "transparent") return true;
    var m = s.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
    if (m) {
      var hex = m[1];
      if (hex.length === 3) {
        hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
      }
      var r = parseInt(hex.slice(0, 2), 16);
      var g = parseInt(hex.slice(2, 4), 16);
      var b = parseInt(hex.slice(4, 6), 16);
      // 相对亮度偏高则视为浅色字
      return (0.299 * r + 0.587 * g + 0.114 * b) > 180;
    }
    var rgb = s.match(/rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)/);
    if (rgb) {
      var rr = +rgb[1], gg = +rgb[2], bb = +rgb[3];
      return (0.299 * rr + 0.587 * gg + 0.114 * bb) > 180;
    }
    return false;
  }

  function scrubLightTextColors(root) {
    if (!root || !root.querySelectorAll) return;
    var nodes = root.querySelectorAll("[style*='color'], font[color]");
    for (var i = 0; i < nodes.length; i++) {
      var el = nodes[i];
      if (el.closest && el.closest("a, .macro-link, .link-internal, #disguise-chrome, #disguise-statusbar, #disguise-panic-sidebar, #disguise-panic-main")) {
        continue;
      }
      try {
        if (el.tagName === "FONT" && el.getAttribute("color")) {
          if (isTooLight(el.getAttribute("color"))) {
            el.removeAttribute("color");
          }
        }
        if (el.style) {
          if (el.style.color && isTooLight(el.style.color)) {
            el.style.color = "#1a1a1a";
          }
          if (el.style.webkitTextFillColor && isTooLight(el.style.webkitTextFillColor)) {
            el.style.webkitTextFillColor = "#1a1a1a";
          }
          // 深色背景在伪装下会变成「色块」，清掉浅色场景下的深底
          if (el.style.backgroundColor) {
            var bg = el.style.backgroundColor.toLowerCase();
            if (bg === "#000" || bg === "#000000" || bg === "black" || bg.indexOf("rgb(0") === 0) {
              el.style.backgroundColor = "transparent";
            }
          }
        }
      } catch (e) { /* ignore */ }
    }
  }

  function ensurePanicLayers() {
    if (!document.body) return;

    // 移除旧版全屏白屏
    var legacy = document.getElementById("disguise-panic");
    if (legacy && legacy.parentNode) legacy.parentNode.removeChild(legacy);

    if (!document.getElementById("disguise-panic-sidebar")) {
      var side = document.createElement("div");
      side.id = "disguise-panic-sidebar";
      side.setAttribute("aria-hidden", "true");
      side.innerHTML =
        '<div class="panic-nav-title">导航窗格 · 学习目录</div>' +
        '<ul class="panic-nav-list">' +
          "<li class=\"active\">前言</li>" +
          "<li>第一章　历史方位</li>" +
          "<li>第二章　指导思想</li>" +
          "<li>第三章　奋斗目标</li>" +
          "<li>第四章　基本方略</li>" +
          "<li>第五章　总体布局</li>" +
          "<li>第六章　战略部署</li>" +
          "<li>第七章　风险挑战</li>" +
          "<li>第八章　党的建设</li>" +
          "<li>结束语</li>" +
        "</ul>" +
        '<div class="panic-nav-foot">内部资料 · 注意保存</div>';
      document.body.appendChild(side);
    }

    if (!document.getElementById("disguise-panic-main")) {
      var main = document.createElement("div");
      main.id = "disguise-panic-main";
      main.setAttribute("aria-hidden", "true");
      main.innerHTML =
        '<div class="panic-paper">' +
          '<div class="panic-paper-head">内部学习资料 · 请勿外传</div>' +
          "<h1>前　言</h1>" +
          '<div class="sub">中共中央宣传部 · 学习读本（节选）</div>' +
          "<p>历史是最好的教科书，也是最好的清醒剂。中国共产党的历史，是一部团结带领全国各族人民进行革命、建设、改革并不断取得辉煌成就的历史。</p>" +
          "<p>学习党史，是坚持和发展中国特色社会主义、把党和国家各项事业继续推向前进的必修课。这门功课不仅必修，而且必须修好。</p>" +
          "<p>本纲要围绕新时代坚持和发展中国特色社会主义的重大时代课题，对党的创新理论作了全面系统阐述，是广大党员干部群众深入学习的重要辅助读物。</p>" +
          "<p>各级党组织要把学习党史同总结经验、观照现实、推动工作结合起来，同解决实际问题结合起来，切实做到学史明理、学史增信、学史崇德、学史力行。</p>" +
          "<p>全党同志要更加紧密地团结在党中央周围，高举中国特色社会主义伟大旗帜，为全面建设社会主义现代化国家、全面推进中华民族伟大复兴而团结奋斗。</p>" +
          '<div class="panic-paper-foot">— 学习贯彻习近平新时代中国特色社会主义思想 —</div>' +
          '<div class="hint">再次按下 ` 或 Esc 返回正文 · Ctrl+Shift+D 退出伪装</div>' +
        "</div>";
      document.body.appendChild(main);
    }
  }

  /** 将掩护层挂到侧栏 / 正文容器内，随布局一起变化 */
  function mountPanicLayers() {
    ensurePanicLayers();
    var side = document.getElementById("disguise-panic-sidebar");
    var main = document.getElementById("disguise-panic-main");
    var bar = document.getElementById("ui-bar");
    var story = document.getElementById("story");

    if (side) {
      var sideHost = bar || document.body;
      if (side.parentNode !== sideHost) sideHost.appendChild(side);
    }
    if (main) {
      var mainHost = story || document.body;
      if (main.parentNode !== mainHost) mainHost.appendChild(main);
    }
  }

  function showPanic() {
    // 若尚未伪装，紧急时先打开伪装（保留 Word 外壳）
    if (!isOn()) setOn(true);
    ensureChrome();
    mountPanicLayers();
    document.documentElement.classList.add("disguise-panic");
    if (document.body) document.body.classList.add("disguise-panic");
    panicVisible = true;
    var side = document.getElementById("disguise-panic-sidebar");
    var main = document.getElementById("disguise-panic-main");
    if (side) side.setAttribute("aria-hidden", "false");
    if (main) main.setAttribute("aria-hidden", "false");
  }

  function hidePanic() {
    document.documentElement.classList.remove("disguise-panic");
    if (document.body) document.body.classList.remove("disguise-panic");
    panicVisible = false;
    var side = document.getElementById("disguise-panic-sidebar");
    var main = document.getElementById("disguise-panic-main");
    if (side) side.setAttribute("aria-hidden", "true");
    if (main) main.setAttribute("aria-hidden", "true");
  }

  function togglePanic() {
    if (panicVisible) hidePanic();
    else showPanic();
  }

  function onKeydown(e) {
    // Ctrl+Shift+D — 主开关
    if (e.ctrlKey && e.shiftKey && (e.key === "D" || e.key === "d")) {
      e.preventDefault();
      e.stopPropagation();
      hidePanic();
      toggle();
      return;
    }

    // ` 紧急掩护（仅侧栏+正文）
    if (e.key === "`" && !e.ctrlKey && !e.altKey && !e.metaKey) {
      // 输入框内不拦截
      var tag = (e.target && e.target.tagName) || "";
      if (tag === "INPUT" || tag === "TEXTAREA" || (e.target && e.target.isContentEditable)) {
        return;
      }
      e.preventDefault();
      togglePanic();
      return;
    }

    // 连按两次 Esc → 紧急掩护
    if (e.key === "Escape") {
      if (panicVisible) {
        e.preventDefault();
        hidePanic();
        lastEscAt = 0;
        return;
      }
      var now = Date.now();
      if (now - lastEscAt < 400) {
        e.preventDefault();
        showPanic();
        lastEscAt = 0;
      } else {
        lastEscAt = now;
      }
    }
  }

  function readInitialState() {
    try {
      var params = new URLSearchParams(window.location.search || "");
      if (params.has("disguise")) {
        return params.get("disguise") !== "0" && params.get("disguise") !== "false";
      }
    } catch (e) { /* ignore */ }

    try {
      return localStorage.getItem(STORAGE_KEY) === "1";
    } catch (e) {
      return false;
    }
  }

  function boot() {
    ensureFab();
    ensureChrome();
    document.addEventListener("keydown", onKeydown, true);

    if (readInitialState()) {
      setOn(true);
    }

    // 段落切换后刷新“页码”并清理浅色字
    try {
      if (window.jQuery) {
        jQuery(document).on(":passagedisplay :passagerender", function () {
          if (!isOn()) return;
          updateStatusbar();
          relocateStatsToStatusbar();
          scrubLightTextColors(document.getElementById("passages"));
          scrubLightTextColors(document.getElementById("ui-bar"));
          scrubLightTextColors(document.getElementById("story-caption"));
          scrubLightTextColors(document.getElementById("storyCaptionDiv"));
          scrubLightTextColors(document.getElementById("storyCaptionContent"));
        });
      }
    } catch (e) { /* ignore */ }

    // 兜底：MutationObserver 监听 passages / 侧栏（#stats 重建）
    var observeRoots = [
      document.getElementById("passages"),
      document.getElementById("story-caption"),
      document.getElementById("ui-bar-body")
    ];
    if (window.MutationObserver) {
      var timer = null;
      var mo = new MutationObserver(function () {
        if (!isOn()) return;
        if (timer) clearTimeout(timer);
        timer = setTimeout(function () {
          updateStatusbar();
          relocateStatsToStatusbar();
          fixNextButtons();
          if (panicVisible) mountPanicLayers();
          scrubLightTextColors(document.getElementById("passages"));
        }, 200);
      });
      for (var i = 0; i < observeRoots.length; i++) {
        if (observeRoots[i]) {
          mo.observe(observeRoots[i], { childList: true, subtree: true, characterData: true });
        }
      }
    }
  }

  // 暴露给控制台手动调用
  window.DisguiseMode = {
    enable: function () { setOn(true); },
    disable: function () { setOn(false); },
    toggle: toggle,
    panic: showPanic,
    isOn: isOn
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }

  // SugarCube 可能较晚 ready，再补一次
  window.addEventListener("load", function () {
    ensureFab();
    if (isOn()) {
      ensureChrome();
      updateStatusbar();
      document.title = TITLE_DISGUISE;
    }
  });

  // 防止游戏脚本改回标题
  try {
    var titleEl = document.querySelector("title");
    if (titleEl && window.MutationObserver) {
      new MutationObserver(function () {
        if (isOn() && document.title !== TITLE_DISGUISE) {
          document.title = TITLE_DISGUISE;
        }
      }).observe(titleEl, { childList: true, characterData: true, subtree: true });
    }
  } catch (e) { /* ignore */ }

  // 周期性兜底（部分环境会直接写 document.title）
  setInterval(function () {
    if (isOn() && document.title !== TITLE_DISGUISE) {
      document.title = TITLE_DISGUISE;
    }
  }, 2000);
})();
