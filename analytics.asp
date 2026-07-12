<%
' Replace this with your real GA4 Measurement ID (example: G-AB12CD34EF)
ga4MeasurementId = "G-3HE08CFG1Z"
%>
<script async src="https://www.googletagmanager.com/gtag/js?id=<%=ga4MeasurementId%>"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());
  gtag('config', '<%=ga4MeasurementId%>');

  function analyticsTrack(eventName, params) {
    if (typeof gtag !== 'function') {
      return;
    }
    gtag('event', eventName, params || {});
  }
</script>