/* global clearInterval, console, setInterval */

let apiKey = null;

function getApiKey() {
  return apiKey || localStorage.getItem("polygonApiKey");
}

/**
 * Gets ticker details from Polygon.io.
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @returns {Promise<string>} JSON string of the ticker details.
 */
async function getTickerDetails(ticker) {
  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error("API key not set. Please set your Polygon.io API key in the task pane.");
  }

  const url = `https://api.polygon.io/v3/reference/tickers?ticker=${ticker}&active=true&limit=100&apiKey=${apiKey}`;

  try {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const data = await response.json();
    return JSON.stringify(data);
  } catch (error) {
    throw new Error(`Failed to fetch ticker details: ${error.message}`);
  }
}


// /**
//  * Add two numbers
//  * @customfunction
//  * @param {number} first First number
//  * @param {number} second Second number
//  * @returns {number} The sum of the two numbers.
//  */
// export function add(first, second) {
//   return first + second;
// }

// /**
//  * Returns "Hello World".
//  * @customfunction
//  * @returns {string} "Hello World"
//  */
// export function helloWorld() {
//   return "Hello World";
// }

// /**
//  * Displays the current time once a second
//  * @customfunction
//  * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
//  */
// export function clock(invocation) {
//   const timer = setInterval(() => {
//     const time = currentTime();
//     invocation.setResult(time);
//   }, 1000);

//   invocation.onCanceled = () => {
//     clearInterval(timer);
//   };
// }

// /**
//  * Returns the current time
//  * @returns {string} String with the current time formatted for the current locale.
//  */
// export function currentTime() {
//   return new Date().toLocaleTimeString();
// }

// /**
//  * Increments a value once a second.
//  * @customfunction
//  * @param {number} incrementBy Amount to increment
//  * @param {CustomFunctions.StreamingInvocation<number>} invocation
//  */
// export function increment(incrementBy, invocation) {
//   let result = 0;
//   const timer = setInterval(() => {
//     result += incrementBy;
//     invocation.setResult(result);
//   }, 1000);

//   invocation.onCanceled = () => {
//     clearInterval(timer);
//   };
// }

// /**
//  * Writes a message to console.log().
//  * @customfunction LOG
//  * @param {string} message String to write.
//  * @returns String to write.
//  */
// export function logMessage(message) {
//   console.log(message);

//   return message;
// }
