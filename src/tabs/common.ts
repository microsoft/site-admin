import { DataSource } from "../ds";

// Returns the storage used info
export function getStorageUsed(): string {
    // See if we can determine the total amount allocated
    if (DataSource.Site.Usage.StoragePercentageUsed > 0) {
        // Return the usage and total allocated
        return `${DataSource.formatBytes(DataSource.Site.Usage.Storage)} of ${DataSource.formatBytes(DataSource.Site.Usage.Storage / DataSource.Site.Usage.StoragePercentageUsed)} (${Math.round(DataSource.Site.Usage.StoragePercentageUsed * 100) + "%"} Used)`;
    } else {
        // Return the usage only
        return DataSource.formatBytes(DataSource.Site.Usage.Storage);
    }
}

// Sees if the properties is empty
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