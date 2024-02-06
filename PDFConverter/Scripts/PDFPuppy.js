var fileArray = {};
var showmergePDFCards = [];
var _PDF_DOC,
    _CURRENT_PAGE,
    _TOTAL_PAGES,
    _PAGE_RENDERING_IN_PROGRESS = 0,
    _CANVAS


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
    $("#imgtxttoPDF").click(function () {
        $("input[id='upldTxttoPDF']").click();
    });
    $("#imgtxttoPDFBulk").click(function () {
        $("input[id='upldTxttoPDFBulk']").click();
    });
    $("#imgupldSplitPDF").click(function () {
        $("input[id='upldSplitPDF']").click();
    });
    $("#imgupldPDFtoWord").click(function () {
        $("input[id='upldPDFtoWord']").click();
    });
    $("#imgupldPDFtoWorddocx").click(function () {
        $("input[id='upldPDFtoWorddocx']").click();
    });

    $("input[id=upldPDFtoWord]").change(function (e) {
        $('.loading').show();
        e.preventDefault();
        var files = $("#upldPDFtoWord").get(0).files;


        var formData = new FormData();

        // Looping over all files and add it to FormData object  
        for (var i = 0; i < files.length; i++) {
            console.log('(files[i].name:' + files[i].name);
            formData.append('product', files[i]);
        }

        $("#upldPDFtoWord")[0].value = '';


        $.ajax({
            type: "POST",
            url: "/PDF/PDFtoWord",
            dataType: "json",
            contentType: false, // Not to set any content header
            processData: false, // Not to process data
            data: formData,
            success: function (result, status, xhr) {
                hideLoader();
                if (result == "Success") {

                    window.location.href = '/PDF/DownloadPDFtoWord';

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
    $("input[id=upldPDFtoWorddocx]").change(function (e) {
        $('.loading').show();
        e.preventDefault();
        var files = $("#upldPDFtoWorddocx").get(0).files;


        var formData = new FormData();

        // Looping over all files and add it to FormData object  
        for (var i = 0; i < files.length; i++) {
            console.log('(files[i].name:' + files[i].name);
            formData.append('product', files[i]);
        }

        $("#upldPDFtoWorddocx")[0].value = '';


        $.ajax({
            type: "POST",
            url: "/PDF/PDFtoWordDocx",
            dataType: "json",
            contentType: false, // Not to set any content header
            processData: false, // Not to process data
            data: formData,
            success: function (result, status, xhr) {
                hideLoader();
                if (result == "Success") {

                    window.location.href = '/PDF/DownloadPDFtoWordDocx';

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


    $("input[id=upldSplitPDF]").change(function (e) {
        $('.loading').show();
        e.preventDefault();
        var files = $("#upldSplitPDF").get(0).files;


        var formData = new FormData();

        // Looping over all files and add it to FormData object  
        for (var i = 0; i < files.length; i++) {
            console.log('(files[i].name:' + files[i].name);
            formData.append('product', files[i]);
        }

        $("#upldSplitPDF")[0].value = '';


        $.ajax({
            type: "POST",
            url: "/PDF/SplitPDF",
            dataType: "json",
            contentType: false, // Not to set any content header
            processData: false, // Not to process data
            data: formData,
            success: function (result, status, xhr) {
                hideLoader();
                if (result == "Success") {

                    window.location.href = '/PDF/DownloadSplitPDFZip';

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

    $("input[id=upldTxttoPDFBulk]").change(function (e) {
        $('.loading').show();
        e.preventDefault();
        var files = $("#upldTxttoPDFBulk").get(0).files;


        var formData = new FormData();

        // Looping over all files and add it to FormData object  
        for (var i = 0; i < files.length; i++) {
            console.log('(files[i].name:' + files[i].name);
            formData.append('product', files[i]);
        }

        $("#upldTxttoPDFBulk")[0].value = '';


        $.ajax({
            type: "POST",
            url: "/PDF/MergeTextToPDFBulk",
            dataType: "json",
            contentType: false, // Not to set any content header
            processData: false, // Not to process data
            data: formData,
            success: function (result, status, xhr) {
                hideLoader();
                if (result == "Success") {

                    window.location.href = '/PDF/DownloadMergetexttoPDFBulk';

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

    $("input[id=upldTxttoPDF]").change(function (e) {
        $('.loading').show();
        e.preventDefault();
        var files = $("#upldTxttoPDF").get(0).files;


        var formData = new FormData();

        // Looping over all files and add it to FormData object  
        for (var i = 0; i < files.length; i++) {
            console.log('(files[i].name:' + files[i].name);
            formData.append('product', files[i]);
        }

        $("#upldTxttoPDF")[0].value = '';


        $.ajax({
            type: "POST",
            url: "/PDF/ConvertWordtoPDF",
            dataType: "json",
            contentType: false, // Not to set any content header
            processData: false, // Not to process data
            data: formData,
            success: function (result, status, xhr) {
                hideLoader();
                if (result == "Success") {

                    window.location.href = '/PDF/DownloadWordtoPDF';

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
            url: "/PDF/ConvertWordtoPDF",           
            dataType: "json",
            contentType: false, // Not to set any content header
            processData: false, // Not to process data
            data: formData,
            success: function (result, status, xhr) {
                hideLoader();
                if (result == "Success") {
                    
                    window.location.href = '/PDF/DownloadWordtoPDF';
                   
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
            url: "/PDF/ConvertImagetoPDF",
            dataType: "json",
            contentType: false, // Not to set any content header
            processData: false, // Not to process data
            data: formData,
            success: function (result, status, xhr) {
                
                //Convert Base64 string to Byte Array.
                hideLoader();
                if (result == "Success") {
                    hideLoader();
                    
                    window.location.href = '/PDF/DownloadImagetoPDF';
                    
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
            url: "/PDF/ConvertImagetoPDF",
            dataType: "json",
            contentType: false, // Not to set any content header
            processData: false, // Not to process data
            data: formData,
            success: function (result, status, xhr) {
                 
                //Convert Base64 string to Byte Array.
                hideLoader();
                if (result == "Success") {
                    
                    window.location.href = '/PDF/DownloadImagetoPDF';
                    
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
    $("input[id=upldMergePDFFiles]").change(function (e) {
        debugger;
        e.preventDefault();
        var temparray=[1,2,3,4,5]
        var length = showmergePDFCards.length;
        var name = '';
         
        if (length < 5) {
            var array3 = temparray.filter(function (obj) { return showmergePDFCards.indexOf(obj) == -1; });

            showmergePDFCards.push(array3[0]);
            var files = $('#upldMergePDFFiles').get(0).files;
            for (var i = 0; i < files.length; i++) {
                name=files[i].name;
                /* formData.append('product', files[i]);*/
                fileArray[array3[0]-1] = files[0];
            }
            var path = URL.createObjectURL(files[0])
            showcard(array3[0], path, name)
        }
        
    });
    $('#submitMergePDF').on('click', function (e) {
        debugger;
        
        
        $('.loading').show();
        e.preventDefault();
        var files = fileArray
        var count=0

        var formData = new FormData();

        // Looping over all files and add it to FormData object  
        //for (var i = 0; i < files.length; i++) {
        //    console.log('(files[i].name:' + files[i].name);
        //    //formData.append('product', files[i]);
            
        //}
        for (var key in files) {
            formData.append('product', files[key]);
            count = count + 1;
        }
        
        $("#upldMergePDFFiles")[0].value = '';
        if (count > 1) {
            $("#mergeModalClose").click();
            //emptyMergeArray();
            $.ajax({
                type: "POST",
                url: "/PDF/MergePDF",
                dataType: "json",
                contentType: false, // Not to set any content header
                processData: false, // Not to process data
                data: formData,
                success: function (result, status, xhr) {

                    //Convert Base64 string to Byte Array.
                    hideLoader();
                    if (result == "Success") {

                        window.location.href = '/PDF/DownloadMergePDF';

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
        }
        else {

            alert('Please  select atleast 2 files!')
        }
    });
   

    $("#MergePDF").on('click', function () {
        showmergePDFCards = [1, 2, 3, 4, 5];
        closecard(showmergePDFCards);
        $('#myModal').show(function () {
          });
        
    });
    $("#mergeModalClose").on('click', function () {
        $('#myModal').hide();
        $('#upldMergePDFFiles').val('');
        hideLoader();
       //emptyMergeArray();
    });

    $("#MergePDFAdd").on('click', function () {
        $("#upldMergePDFFiles")[0].value = '';
        $("input[id='upldMergePDFFiles']").click();

    });

    $('#closeicon1').on('click', function () {
        debugger;
        removeCards(1);
        $(this).closest('.card').fadeOut();
        delete fileArray[0];
    })
    $('#closeicon2').on('click', function () {
        debugger;
        removeCards(2);
        $(this).closest('.card').fadeOut();
        delete fileArray[1];
    })
    $('#closeicon3').on('click', function () {
        debugger;
        removeCards(3);
        $(this).closest('.card').fadeOut();
        delete fileArray[2];
    })
    $('#closeicon4').on('click', function () {
        debugger;
        removeCards(4);
        $(this).closest('.card').fadeOut();
        delete fileArray[3];
    })
    $('#closeicon5').on('click', function () {
        debugger;
        removeCards(5);
        $(this).closest('.card').fadeOut();
        delete fileArray[4];
    })
})
function closecard(array) {

    for (i = 0; i < array.length; i++) {
        switch (showmergePDFCards[i]) {
            case 1: $('#closeicon1').click();
            case 2: $('#closeicon2').click();
            case 3: $('#closeicon3').click();
            case 4: $('#closeicon4').click();
            case 5: $('#closeicon5').click();
        }
    }
}

function showcard(array,path,name) {

    


    switch (array) {
        case 1: { $('#card1').show(function () { showPDF(path, 'pdf-canvas1'); }) } $('#canvas1name').empty(); $('#canvas1name').html(name); break;
        case 2: { $('#card2').show(function () { showPDF(path, 'pdf-canvas2'); }) } $('#canvas2name').empty(); $('#canvas2name').html(name); break;
        case 3: { $('#card3').show(function () { showPDF(path, 'pdf-canvas3'); }) } $('#canvas3name').empty(); $('#canvas3name').html(name); break;
        case 4: { $('#card4').show(function () { showPDF(path, 'pdf-canvas4'); }) } $('#canvas4name').empty(); $('#canvas4name').html(name); break;
        case 5: { $('#card5').show(function () { showPDF(path, 'pdf-canvas5'); }) } $('#canvas5name').empty(); $('#canvas5name').html(name); break;
    }
    
}
function removeCards(i) {
    const index = showmergePDFCards.indexOf(i);
    if (index > -1) { // only splice array when item is found
        showmergePDFCards.splice(index, i); // 2nd parameter means remove one item only
    }

}
function hideLoader() {

    $('.loading').fadeOut(2000)

}

function erroralert() {


    setTimeout(function () {

        alert('Please check the file format and size and try one more time!')
    }, 2000);



}


//function emptyMergeArray() {
//    showmergePDFCards = [];
//    fileArray = {};
//}
async function showPDF(pdf_url, pdfcanvas) {
    _PDF_DOC = '';
    //document.querySelector("#pdf-loader").style.display = 'block';
     _CANVAS = document.getElementById(("#" + pdfcanvas));
    // get handle of pdf document
    try {
        _PDF_DOC = await pdfjsLib.getDocument({ url: pdf_url });
    }
    catch (error) {
        alert(error.message);
    }

    // total pages in pdf
     _TOTAL_PAGES = _PDF_DOC.numPages;

    // Hide the pdf loader and show pdf container
    //document.querySelector("#pdf-loader").style.display = 'none';
    //document.querySelector("#pdf-contents").style.display = 'block';

    // show the first page
     showPage(1, pdfcanvas);
    document.querySelector("#pdf-contents").style.display = 'block';
}


async function showPage(page_no, pdfcanvas) {
     _PAGE_RENDERING_IN_PROGRESS = 1;
     _CURRENT_PAGE = page_no;

    _CANVAS = document.querySelector('#' + pdfcanvas);
    // while page is being rendered hide the canvas and show a loading message
    document.querySelector("#" + pdfcanvas).style.display = 'none';
    // get handle of page
    try {
        var page =  await  _PDF_DOC.getPage(page_no);
    }
    catch (error) {
        //alert(error.message);
    }

    // original width of the pdf page at scale 1
    var pdf_original_width = page.getViewport(1).width;

    // as the canvas is of a fixed width we need to adjust the scale of the viewport where page is rendered
    var scale_required = _CANVAS.width / pdf_original_width;

    // get viewport to render the page at required scale
    var viewport = page.getViewport(scale_required);

    // set canvas height same as viewport height
    _CANVAS.height = viewport.height;

    // setting page loader height for smooth experience
    // page is rendered on <canvas> element
    var render_context = {
        canvasContext:  _CANVAS.getContext('2d'),
        viewport: viewport
    };

    // render the page contents in the canvas
    try {
        await page.render(render_context);
    }
    catch (error) {
        alert(error.message);
    }

    _PAGE_RENDERING_IN_PROGRESS = 0;

    // show the canvas and hide the page loader
    
    document.querySelector("#"+pdfcanvas).style.display = 'block';
   
}





