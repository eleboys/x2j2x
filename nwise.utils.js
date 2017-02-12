
var NWISE = {};
(function make_nwise(NWISE) {

    function join_object_path(p1,p2) {
        if (!p1) {
            return p2;
        }
        if (!p2){
            return p1;
        }
        return `${p1}.${p2}`;
    }

    function flatten_object(jsonData, basePath, output) {
        var output = output || [];
        basePath = basePath || '';
        if (!Object.isExtensible(jsonData)){
            output.push([ basePath,  jsonData]);
            return output;          
        }
        for (var key in jsonData) {
            var value = jsonData[key];
            if (Object.isExtensible(value)) {
                if (!Array.isArray(value)) {
                    flatten_object(value, join_object_path(basePath, key), output);
                }else {
                    value.forEach((item,i) =>{
                        flatten_object(item,  join_object_path(basePath, key)+`[${i}]`, output);
                    });
                }
            }else {
                output.push([join_object_path(basePath, key), value]);
            }
        }

        return output;
    }
    var utils = {
        flatten_object: flatten_object
    }
    NWISE.utils = utils;
})(typeof exports !== 'undefined' ? exports : NWISE);