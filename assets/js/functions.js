$( document ).ready(function() {

  // Get started!
    
    $(".overlay").on("click",function(){
       var link = $(this).prev().attr('src');
        $(".img-open .content img").attr('src',link);
       $(".img-open").addClass("opened");
       
        
    });
    
    $(".close-btn").on("click",function(){
       $(".img-open").removeClass("opened"); 
    });
    
});
