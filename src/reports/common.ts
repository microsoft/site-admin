export function sanitizeString(value: string): string {
    // Return an empty string if the value is null or undefined
    if (value == null) { return ""; }

    // Replace the < and > w/ &lt; &gt;
    return value.replace(/[<>]/g, (match) => {
        if (match === "<") return "&lt;";
        if (match === ">") return "&gt;";
        return match;
    });
}