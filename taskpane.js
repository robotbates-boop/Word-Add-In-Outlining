(function () {
  function setStatus(msg) {
    var el = document.getElementById("status");
    if (el) el.textContent = msg;
  }

  function load(src) {
    return new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = src + "?v=" + Date.now(); // always fresh
      s.async = true;
      s.onload = function () { resolve(); };
      s.onerror = function () { reject(new Error("Failed to load: " + src)); };
      document.head.appendChild(s);
    });
  }

  setStatus("Bootstrap loaded. Loading main…");

  load("taskpane_main.js")
    .then(function () {})
    .catch(function (e) { setStatus("ERROR: " + (e && e.message ? e.message : String(e))); });
})();
