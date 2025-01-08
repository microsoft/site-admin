export function isEmpty(obj) {
    // Return true if null
    if (obj == null) { return true; }

    // Parse the properties
    for (var prop in obj) {
        // See if the object has properties
        if (Object.prototype.hasOwnProperty.call(obj, prop)) { return false; }
    }

    // Object is empty
    return true;
}