$(document).ready(function () {
    //deletefiles();
    $('.loading').hide(); 
    setInterval
    $("#imgWordtoPDF").click(function () {
        $("input[id='upldWordtoPDF']").click();
    });
    $("#imgjpegtoPDF").click(function () {
        $("input[id='upldjpegtoPDF']").click();
    });
    $("#imgpngtoPDF").click(function () {
        $("input[id='upldpngtoPDF']").click();
    });
    $("#tiffAndgifftoPDF").click(function () {
        $("input[id='upldtiffAndgifftoPDF']").click();
    });
    /*$(':upldWordtoPDF').on('click', function () { alert('click') })*/
    $("input[id=upldWordtoPDF]").change(function (e) {
          $('.loading').show();
        e.preventDefault();
        var files = $("#upldWordtoPDF").get(0).files;


        var formData = new FormData();

        // Looping over all files and add it to FormData object  
        for (var i = 0; i < files.length; i++) {
            console.log('(files[i].name:' + files[i].name);
            formData.append('product', files[i]);
        }

        $("#upldWordtoPDF")[0].value = '';


        $.ajax({
            type: "POST",
            url: "/Home/ConvertWordtoPDF",
            dataType: "json",
            contentType: false, // Not to set any content header
            processData: false, // Not to process data
            data: formData,
            success: function (result, status, xhr) {
                hideLoader();
                if (result == "Success") {
                    
                    window.location.href = '/Home/DownloadWordtoPDF';
                   
                }
                else {

                    erroralert();
                }
            },
            error: function (xhr, status, error) {
                
                hideLoader();                
                erroralert();


            }
        });
    });



    /*$(':upldWordtoPDF').on('click', function () { alert('click') })*/
    $("input[id=upldjpegtoPDF]").change(function (e) {
         
        $('.loading').show();
        e.preventDefault();
        var files = $("#upldjpegtoPDF").get(0).files;


        var formData = new FormData();

        // Looping over all files and add it to FormData object  
        for (var i = 0; i < files.length; i++) {
            console.log('(files[i].name:' + files[i].name);
            formData.append('product', files[i]);
        }


        $("#upldjpegtoPDF")[0].value = '';

        $.ajax({
            type: "POST",
            url: "/Home/ConvertImagetoPDF",
            dataType: "json",
            contentType: false, // Not to set any content header
            processData: false, // Not to process data
            data: formData,
            success: function (result, status, xhr) {
                
                //Convert Base64 string to Byte Array.
                hideLoader();
                if (result == "Success") {
                    hideLoader();
                    
                    window.location.href = '/Home/DownloadImagetoPDF';
                    
                }
                else {

                    erroralert();
                }
            }

            ,
            error: function (xhr, status, error) {
                hideLoader();
                erroralert();

            }
        });
    });

    $("input[id=upldpngtoPDF]").change(function (e) {
          $('.loading').show();
        e.preventDefault();
        var files = $("#upldpngtoPDF").get(0).files;


        var formData = new FormData();

        // Looping over all files and add it to FormData object  
        for (var i = 0; i < files.length; i++) {
            console.log('(files[i].name:' + files[i].name);
            formData.append('product', files[i]);
        }

        $("#upldpngtoPDF")[0].value = '';


        $.ajax({
            type: "POST",
            url: "/Home/ConvertImagetoPDF",
            dataType: "json",
            contentType: false, // Not to set any content header
            processData: false, // Not to process data
            data: formData,
            success: function (result, status, xhr) {
                 
                //Convert Base64 string to Byte Array.
                hideLoader();
                if (result == "Success") {
                    
                    window.location.href = '/Home/DownloadImagetoPDF';
                    
                }
                else {

                    erroralert();
                }
            }

            ,
            error: function (xhr, status, error) {
                hideLoader();
                erroralert();

            }
        });

    });
    $("input[id=upldtiffAndgifftoPDF]").change(function (e) {

        $('.loading').show();
        e.preventDefault();
        var files = $("#upldtiffAndgifftoPDF").get(0).files;


        var formData = new FormData();

        // Looping over all files and add it to FormData object  
        for (var i = 0; i < files.length; i++) {
            console.log('(files[i].name:' + files[i].name);
            formData.append('product', files[i]);
        }


        $("#upldtiffAndgifftoPDF")[0].value = '';

        $.ajax({
            type: "POST",
            url: "/Home/ConvertImagetoPDF",
            dataType: "json",
            contentType: false, // Not to set any content header
            processData: false, // Not to process data
            data: formData,
            success: function (result, status, xhr) {

                //Convert Base64 string to Byte Array.
                hideLoader();
                if (result == "Success") {
                    hideLoader();

                    window.location.href = '/Home/DownloadImagetoPDF';

                }
                else {

                    erroralert();
                }
            }

            ,
            error: function (xhr, status, error) {
                hideLoader();
                erroralert();

            }
        });
    });




})


function hideLoader() {

    $('.loading').fadeOut(2000)

}

function erroralert() {


    setTimeout(function () {

        alert('Please check the file format and size and try one more time!')
    }, 2000);



}

