function showFullMoveHideHalfMove () {
    const halfMoveNodeList = document.querySelectorAll(".halfmove")
    const fullMoveNodeList = document.querySelectorAll(".fullmove")
    for( const node of halfMoveNodeList) {
        node.style.display = "none";
    }
    for( const node of fullMoveNodeList) {
        node.style.display = "inline";
    }
}
function showHalfMoveHideFullMove () {
    const halfMoveNodeList = document.querySelectorAll(".halfmove")
    const fullMoveNodeList = document.querySelectorAll(".fullmove")
    for( const node of halfMoveNodeList) {
        node.style.display = "inline";
    }
    for( const node of fullMoveNodeList) {
        node.style.display = "none";
    }
}
function openTab(evt, tabName) {
   var i, tabcontent, tablinks;
   tabcontent = document.getElementsByClassName("tab");
   for (i = 0; i < tabcontent.length; i++) {
     tabcontent[i].style.display = "none";
   }
   tablinks = document.getElementsByClassName("tablinks");
   document.getElementById(tabName).style.display = "grid";
   if( evt != null) {
     for (i = 0; i < tablinks.length; i++) {
       tablinks[i].className = tablinks[i].className.replace(" active", "");
     }
     evt.currentTarget.className += " active";
   }
} 
var tabcontent = document.getElementsByClassName("tab");
for (i = 1; i < tabcontent.length; i++) {
  tabcontent[i].style.display = "none";
}

