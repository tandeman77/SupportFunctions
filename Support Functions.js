function indexOf2d(array, find, col){
    var i = 0;
    for (i=0;i<array.length;i++){
      if(array[i][col] == find){
        return i;
      }    
    }
    return -1;
  }

//======================================================================================================================
//how ot use:  keywordList.sort(sortFunction);
function sortFunctiona2z(a, b) {
    if (a[0] === b[0]) {
        return 0;
    }
    else {
        return (a[0] < b[0]) ? -1 : 1;
    }
}

//======================================================================================================================
//how ot use:  keywordList.sort(sortFunction);
function sortFunctionz2a(a, b) {
    if (a[0] === b[0]) {
        return 0;
    }
    else {
        return (a[0] > b[0]) ? -1 : 1;
    }
}

//======================================================================================================================
// how to use: keywordList.sort(compareSecondColumnDescending);
function compareSecondColumnDescending(a, b) {
    if (a[1] === b[1]) {
        return 0;
    }
    else {
        return (a[1] > b[1]) ? -1 : 1;
    }
}

//======================================================================================================================
// how to use: keywordList.sort(compareSecondColumnAscending);
function compareSecondColumnAscending(a, b) {
if (a[1] === b[1]) {
    return 0;
}
else {
    return (a[1] < b[1]) ? -1 : 1;
}
}