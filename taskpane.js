(function () {
  function load(src) {
    var s = document.createElement("script");
    s.src = src + "?v=" + Date.now();
    s.async = true;
    s.onerror = function () {
      var el = document.getElementById("status");
      if (el) el.textContent = "ERROR: failed to load " + src;
    };
    document.head.appendChild(s);
  }

  // Always load the real logic fresh
  load("taskpane_main.js");
})();
