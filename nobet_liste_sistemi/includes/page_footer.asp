<!-- #include file="../lib/encoding.asp" -->
<%
Call NobetIstekBitir()
%>
<script>
(function () {
  document.documentElement.classList.add("page-ready");
  window.addEventListener("pageshow", function () {
    document.body.classList.remove("page-leave");
  });

  var links = document.querySelectorAll("a[href]");
  for (var i = 0; i < links.length; i++) {
    (function (link) {
      if (link.classList.contains("home-icon-btn") || link.classList.contains("no-transition")) {
        return;
      }
      var href = link.getAttribute("href") || "";
      if (!href || href.charAt(0) === "#" || link.target === "_blank" || link.hasAttribute("download")) {
        return;
      }
      if (href.indexOf("portal_restore.asp") >= 0 || href.indexOf("10.201.65.10") >= 0) {
        return;
      }
      if (href.indexOf("javascript:") === 0 || href.indexOf("mailto:") === 0) {
        return;
      }
      link.addEventListener("click", function (event) {
        if (event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) {
          return;
        }
        document.body.classList.add("page-leave");
      });
    })(links[i]);
  }
})();
</script>
