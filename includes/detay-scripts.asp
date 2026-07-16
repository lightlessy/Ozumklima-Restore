<script src="js/darkbox.js"></script>
<script>
$(function () {
    $('.product-thumb').on('click', function () {
        var button = $(this);
        var imagePath = button.attr('data-image');

        $('.product-thumb').removeClass('is-active');
        button.addClass('is-active');
        $('#productMainImage').attr('src', imagePath);
        $('#productMainLink').attr('href', imagePath).attr('data-darkbox', imagePath);
    });

    $('.product-color-dot[data-color]').each(function () {
        var colorValue = $(this).attr('data-color');
        if (/^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(colorValue)) {
            this.style.backgroundColor = colorValue;
        }
    });

    $('.js-product-interest').on('click', function () {
        if (typeof analyticsTrack === 'function') {
            analyticsTrack('product_interest_click', {
                cta: $(this).attr('data-cta'),
                product_id: $(this).attr('data-product-id')
            });
        }
    });
});
</script>
<%
DpCloseRecordset relatedProducts
DpCloseConnection dpConn
%>
</body>
</html>
