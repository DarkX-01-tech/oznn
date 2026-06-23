<!-- #include file="../lib/encoding.asp" -->
<%
Call PortalOturumKodSayfasiSifirla()
%>
<script>
(function () {
  document.documentElement.classList.add("page-ready");

  var links = document.querySelectorAll("a[href]");
  for (var i = 0; i < links.length; i++) {
  (function (link) {
    var href = link.getAttribute("href");
    if (!href || href.charAt(0) === "#" || link.target === "_blank" || link.hasAttribute("download")) {
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
