/** Collection of helpful functions that don't fit anywhere else */



/** 
 * search through a sorted an array sorted by col, for a specified key.
 * returns value
 * returns -1 if not found
 */
function binarySearch(arr, key, index, low, high) {
    var mid = parseInt((high + low) / 2);

    if (arr[mid][0] == key) {
        return arr[mid][index]
    } else if (low >= high) {
        return -1
    } else if (arr[mid][0] > key) {
        return binarySearch(arr, key, index, low, mid - 1)
    } else {
        return binarySearch(arr, key, index, mid + 1, high)
    }
}

/*
 * indexOf() - To find the index of a specific key in the column of a 2D array
 * Input
 * arr: array to search
 * key: the value to find
 * col: which column to search
 * low/high: range in which to search for the key
 * 
 * output: index position of the key
 * 
 * example: indexOf([a,b,c,f,h,j],a,0,0,5) returns 0.
 * returns -1 if not found
 */
function indexOf(arr, key, col, low, high) {
    var mid = parseInt((high + low) / 2);

    if (arr[mid][col] == key) {
        return mid
    } else if (low >= high) {
        return -1
    } else if (arr[mid][col] > key) {
        return indexOf(arr, key, col, low, mid - 1)
    } else {
        return indexOf(arr, key, col, mid + 1, high)
    }
}

/*
* Function mergeArrays
* Smears an array a b into an array a.
* inputs: arrays a, b
* returns: array c
* 
* example:
* mergeArrays( [1,2,3,4],[a, ,b, ,c] )
* gives the result
* [a, 1, b, 2, c, 3, 4]
* 
*/
function mergeArrays(aArr, bArr) {

    var row = 0,
        bLen = bArr.length,
        aLen = aArr.length;
    var result = new Array();

    if (aLen == 0) {
        return bArr
    }

    for (var i = 0; i < (aLen); i++) {

        // find first vacant row to throw the ID in
        while (bArr[row] != "" && row < bArr.length) {
            result[row] = bArr[row]
            row++
        }

        // Vacant row found. put an ID from IDs in the empty space
        result[row] = aArr[i]
        row++
    }

    return result
}

/**
 * 
 * sort a 2D array by column k 
 * 
 */
function sortBy(k) {
    return function (a, b) {
        if (a[k] === b[k]) {
            return 0;
        } else {
            return (a[k] < b[k]) ? -1 : 1;
        }
    }

}