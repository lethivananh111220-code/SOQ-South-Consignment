const targetTimestamp = new Date(2026, 5, 13).getTime(); // 13-Thg6
let possibleNextDeliveryTimestamps = [
    new Date(2026, 5, 8).getTime(),
    new Date(2026, 5, 9).getTime(),
    new Date(2026, 5, 10).getTime(),
    new Date(2026, 5, 11).getTime(),
    new Date(2026, 5, 12).getTime(),
    new Date(2026, 5, 13).getTime(),
    new Date(2026, 5, 14).getTime(),
    new Date(2026, 5, 15).getTime()
];
let futureDates = possibleNextDeliveryTimestamps.filter(t => t > targetTimestamp + 3600000);
let dynamicLT = 0;
if (futureDates.length > 0) {
    let nextTS = Math.min(...futureDates);
    dynamicLT = Math.round((nextTS - targetTimestamp) / 86400000);
}
console.log('dynamicLT: ' + dynamicLT);
