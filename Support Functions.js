function indexOf2d(array, find, col) {
    var i = 0;
    for (i = 0; i < array.length; i++) {
        if (array[i][col] == find) {
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

function arrayTo1D(arr) {
    var output = [];
    for (var i in arr) {
        output.push(arr[i][0]);
    };
    return output;
};
//=======================================================================================
function strReplace(str, replaceText, replaceWith) {
    //replace a string with whatever
    var text = str.toString();
    if (replaceText.length > 1) {
        var i = 0;
        for (i = 0; i < replaceText.length; i++) {
            text = text.replace(replaceText[i], replaceWith);
        };
    } else {
        text = text.replace(replaceText, replaceWith);
    };
    return text
}
//=======================================================================================
function cleanArr(arr, textToReplace, replaceTextWith, dimension) {
    //function to clean an array of a value or an array of values
    //have to use in conjunction of the strReplace function
    //dimension is optional. only needed if work with multi-dimensional array
    var output = [];
    if (arguments.length == 3) {
        for (var i = 0; i < arr.length; i++) {
            output.push(strReplace(arr[i], textToReplace, replaceTextWith));
        }
    } else if (arguments.length == 4) {
        for (var i = 0; i < arr.length; i++) {
            output.push(strReplace(arr[i][dimension], textToReplace, replaceTextWith));
        }
    }
    return output;
}
//=======================================================================================
function getCurrentAccountDetails() {
    var currentAccount = AdWordsApp.currentAccount();
    Logger.log('Customer ID: ' + currentAccount.getCustomerId() +
        ', Currency Code: ' + currentAccount.getCurrencyCode() +
        ', Timezone: ' + currentAccount.getTimeZone());
    var stats = currentAccount.getStatsFor('LAST_MONTH');
    Logger.log(stats.getClicks() + ' clicks, ' +
        stats.getImpressions() + ' impressions last month');
}
//=======================================================================================
function turnArrTo2d(arr) {
    var i = 0
    for (i = 0; i < arr.length; i++) {
        arr[i] = [arr[i]];
    }
    return arr
}
//=======================================================================================
function turnArrTo1d(arr) {
    var i = 0
    var resultArr = [];
    for (i = 0; i < arr.length; i++) {
        resultArr[i] = arr[i][0];
    }
    return resultArr
}
//=======================================================================================
function getUnique(arr, dimension) {
    //go through an array and extract only unique values from that array
    //dimension is optional
    var outputArray = []
    if (arguments.length == 1) {
        arr.forEach(function (arrValue) {
            //not sure if a value is not already in an array, what would be the output of indexOf.
            if (outputArray.indexOf(arrValue) !== undefined) {
                outputArray.push(arrValue);
            };
        });
    } else if (arguments.length == 2) {
        arr[dimension].forEach(function (arrValue) {
            //not sure if a value is not already in an array, what would be the output of indexOf.
            if (outputArray.indexOf(arrValue) !== undefined) {
                outputArray.push(arrValue);
            };
        });
    }
    return outputArray;
}
//=======================================================================================
function deleteEmptyInArray(arr) {
    //remove position in an array that's empty
    var output = [];

    return output
}
//=======================================================================================
function stringSimilarity(sa1, sa2) {
    // Compare two strings to see how similar they are.
    // Answer is returned as a value from 0 - 1
    // 1 indicates a perfect similarity (100%) while 0 indicates no similarity (0%)
    // Algorithm is set up to closely mimic the mathematical formula from
    // the article describing the algorithm, for clarity.
    // Algorithm source site: http://www.catalysoft.com/articles/StrikeAMatch.html
    // (Most specifically the slightly cryptic variable names were written as such
    // to mirror the mathematical implementation on the source site)
    //
    // 2014-04-03
    // Found out that the algorithm is an implementation of the Sørensen–Dice coefficient [1]
    // [1] http://en.wikipedia.org/wiki/S%C3%B8rensen%E2%80%93Dice_coefficient
    //
    // The algorithm is an n-gram comparison of bigrams of characters in a string


    // for my purposes, comparison should not check case or whitespace
    var s1 = sa1.replace(/\s/g, "").toLowerCase();
    var s2 = sa2.replace(/\s/g, "").toLowerCase();

    function intersect(arr1, arr2) {
        // I didn't write this.  I'd like to come back sometime
        // and write my own intersection algorithm.  This one seems
        // clean and fast, though.  Going to try to find out where
        // I got it for attribution.  Not sure right now.
        var r = [], o = {}, l = arr2.length, i, v;
        for (i = 0; i < l; i++) {
            o[arr2[i]] = true;
        }
        l = arr1.length;
        for (i = 0; i < l; i++) {
            v = arr1[i];
            if (v in o) {
                r.push(v);
            }
        }
        return r;
    }

    var pairs = function (s) {
        // Get an array of all pairs of adjacent letters in a string
        var pairs = [];
        for (var i = 0; i < s.length - 1; i++) {
            pairs[i] = s.slice(i, i + 2);
        }
        return pairs;
    }

    var similarity_num = 2 * intersect(pairs(s1), pairs(s2)).length;
    var similarity_den = pairs(s1).length + pairs(s2).length;
    var similarity = similarity_num / similarity_den;
    return similarity;
};
//=======================================================================================
function averageInArr(arr) {
    //average all numbers in array;
    var averageOutput = 0;
    for (var i = 0; i < arr.length; i++) {
        averageOutput += arr[i];
    }

    return averageOutput / arr.length;
}
//=======================================================================================
function bestMatch(camAd, existingKeywords, searchQueries, camAdIndex) {
    /*find the best match for a string.
    arr should be the keywords wtihin the account
    test should be the search query
    compareDimension = the dimension of arr array to use in the stringsimilarity formula.
    returnDimension is the dimension of the inputArr that should be the output.
    this should return string
    */
    //=======================================================================================
    //get the array ready to be processed in the next step and get the best adGroup.
    var output = [];
    var test = camAd.length
    for (var i = 0; i < camAd.length; i++) {
        output.push([camAdIndex[i], camAd[i], existingKeywords[i], stringSimilarity(existingKeywords[i], searchQueries)]);
    }

    //get the best ad group
    //up to here. best match is not correct.
    var adgroupScore = [];
    var averageScore = 0;
    averageScore = output[0][3];
    var i = 0
    for (i = 1; i < output.length; i++) {
        /*
        go through all the output array
        if number is first of its kind
        then AverageScore = output[i][3];
        if number is the same as previous then averageScore = (averageScore+output[i][3])/2
        then record score
        adgroupScore.push([output[i][1],averageScore]);
        */
        if (output[i][0] == output[i - 1][0]) {
            averageScore = (averageScore + output[i][3]) / 2
        } else {
            adgroupScore.push([output[i - 1][1], averageScore]);
            averageScore = output[i][3];
        }
    }
    adgroupScore = sortMultiArray(adgroupScore, 2, 1)
    return adgroupScore[0][0];
}
//=======================================================================================
function multiArrayTo1d(arr, dimension) {
    // process multiple dimensional array and give one dimensional array
    // array much be in form of [] []
    var output = [];
    for (var i = 0; i < arr.length; i++) {
        output.push(arr[i][dimension]);
    }
    return output;
}
//=======================================================================================
function replaceTextFromArray(arr, arrOfText, replaceWith) {
    //look at an array, replace keywords from an array with replaceWith text
    //arr must be 1d array
    var output = [];
    arr.forEach(function (arrValue) {
        for (var i = 0; i < arrofText.length; i++) {
            output.push(replace(arr, arrOfText[i], ''));
        }
    });
    return output;
}
//=======================================================================================
function writeToSheet(sheet, range, arr) {
    //this function write an array value to a sheet in a specified starting range e.g. A1
    //array mush be a 1d array
    //check formula for numberValue, replace, left and right in JS.
    var output = [];
    arr.forEach(function (arrValue) {
        output.push([arrayValue]);
    })
    var outputRange = '';
    outputRange = range & ':' & range.substring(0, 1) & (number(range.replace(range.substring(0, 1))) + arr.length);
    sheet.getRange(outputRange).setValues(output);
}
//=======================================================================================
function backToSheetReady(arr) {
    //turn a 1d array into an array ready to setValues
    var output = [];
    arr.forEach(function (arrValue) {
        output.push([arrValue]);
    })
    return output;
}

//=======================================================================================
function sortMultiArray(arr, dimension, sortOrder) {
    //sort multi dimensional array
    //dimension 2 = 2d array, 3 = 3d array, etc.
    //only works with integer
    //sort order; 0 = ascending, 1 = descending

    //sort by first value in the array
    arr = arr.sort(function (a, b) {
        return a[dimension - 1] - b[dimension - 1];
    });
    if (sortOrder == 0) {
        return arr;
    } else if (sortOrder == 1) {
        return arr.reverse();
    }
}
