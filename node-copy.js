var inObject = {
    child: [
        {
            child: [
                {
                    child: [
                        {
                            value: "mystring1"
                        }
                    ]
                },
                {
                    child: [
                        {
                            value: "mystring2"
                        }
                    ]
                }
            ]
        },
        {
            child: [
                {
                    value: "mystring3"
                }
            ]
        }
    ]
};

var result = makeObjects(inObject);


function makeObjects(from) {
    if (from.child) {
        return makeValue(from.value);
    } else {
        return makeChild(from.child);
    }
}
function makeValue(value) {
    var to = new Object;

    to.value = value;

    return to;
}
function makeChild(child) {
    var to = new Object;

    to.child = new Array;
    for (var offset = 0; offset < child.length; offset++) {
        to.child[offset] = makeObjects(child[offset]);
    }

    return to;
}
