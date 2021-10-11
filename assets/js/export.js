(function($){
    $(".export-data").on('click', function () {
        var plugin = $('select[name="plugin"]').val();
        var nonce = $('#nonce').val();
        var loading = $('.loading .status');
        loading.removeClass('error').addClass('run').html('<span>Fetch data from mail</span><img class="loading-image" src="assets/images/ajax-loader.gif"/>');
        $.ajax({
            url : 'fetchdata.php',
            type : 'POST',
            dataType : 'json',
            data: {
                'plugin' : plugin,
                'nonce' : nonce
            },
            success : function (res) {
                if (res === -1) {
                    loading.addClass('error').html('An error occurred');
                    return;
                }

                if (res.status === false) {
                    loading.addClass('error').html(res.message);
                    return;
                }
                // Make file excel
                loading.removeClass('error').addClass('run').html('<span>Initialize excel file</span><img class="loading-image" src="assets/images/ajax-loader.gif" />');

                setTimeout(function(){
                    $.ajax({
                        url : 'makexlsx.php',
                        type : 'POST',
                        dataType : 'json',
                        data: {
                            'nonce' : nonce,
                            'datas' : res.message,
                        },
                        success : function (data) {
                            if (data === -1) {
                                loading.addClass('error').html('An error occurred');
                                return;
                            }

                            if (data.status === false) {
                                loading.addClass('error').html(data.file);
                                return;
                            }

                            var time = new Date().toISOString().slice(0,10);
                            var $a = $("<a>");
                            $a.attr("href",data.file);
                            $("body").append($a);
                            $a.attr("download", plugin + "-feedback-" + time + ".xlsx");
                            $a[0].click();
                            $a.remove();

                            // Message done!
                            loading.removeClass('error').addClass('run').html('Done!');
                            setTimeout(function() {
                                loading.removeClass('run error').html('Not available');
                            }, 2000);
                        }
                    })
                } , 2000);
            }
        })
    });
}(jQuery));
