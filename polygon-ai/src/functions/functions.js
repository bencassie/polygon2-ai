/* global clearInterval, console, setInterval */

let apiKey = null;

function getApiKey() {
  return apiKey || localStorage.getItem("polygonApiKey");
}

/**
 * Gets a specific ticker detail from Polygon.io.
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @param {string} [property] The property to retrieve (e.g., "name", "market"). If omitted, returns a formatted string.
 * @returns {Promise<string>} The value of the specified property or a formatted string if no property is provided.
 */
export async function getTickerDetails(ticker, property) {
  try {
    const apiKey = getApiKey(); // Assume this retrieves your Polygon.io API key
    if (!apiKey) {
      return "API key not set. Please set your Polygon.io API key.";
    }

    const url = `https://api.polygon.io/v3/reference/tickers?ticker=${ticker}&active=true&limit=100&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return `HTTP error! status: ${response.status}`;
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      const details = data.results[0]; // The ticker details object

      if (property) {
        // Return the specific property if provided
        return details[property] || "Property not found.";
      } else {
        // Return a formatted string if no property is specified
        return `${details.ticker} - ${details.name} (${details.market})`;
      }
    } else {
      return "No ticker details found.";
    }
  } catch (error) {
    return `Error: ${error.message}`;
  }
}

/**
 * Retrieves news articles for a specific stock ticker from Polygon.io.
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @param {number} [limit=10] The maximum number of news articles to retrieve.
 * @returns {Promise<string[][]>} A 2D array of news articles with title, description, and URL.
 */
export async function getTickerNews(ticker, limit = 10) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      throw new Error("API key not set. Please set your Polygon.io API key.");
    }
    limit = (limit && limit > 0) ? limit : 10;
    const url = `https://api.polygon.io/v2/reference/news?ticker=${ticker}&limit=${limit}&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      const newsArticles = data.results.map(article => [
        article.title || "N/A",
        article.description || "N/A",
        article.article_url || "N/A"
      ]);
      return newsArticles;
    } else {
      return [["No news articles found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves the latest stock price data for a ticker from Polygon.io.
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @param {string} [property] Optional property to return (e.g., "close", "high", "low", "open", "volume").
 * @returns {Promise<number|string>} The latest price data or specified property.
 */
export async function getLatestPrice(ticker, property) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return "API key not set. Please set your Polygon.io API key.";
    }

    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/prev?adjusted=true&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return `HTTP error! status: ${response.status}`;
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      const priceData = data.results[0];
      
      if (property && priceData[property] !== undefined) {
        return priceData[property];
      } else if (property) {
        return "Property not found.";
      } else {
        return priceData.c; // Default to close price if no property specified
      }
    } else {
      return "No price data found.";
    }
  } catch (error) {
    return `Error: ${error.message}`;
  }
}

/**
 * Gets the market status from Polygon.io.
 * @customfunction
 * @param {string} [market="us"] The market to check (e.g., "us", "fx", "crypto").
 * @returns {Promise<string>} The current market status.
 */
export async function getMarketStatus(market = "us") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return "API key not set. Please set your Polygon.io API key.";
    }

    const url = `https://api.polygon.io/v1/marketstatus/now?apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return `HTTP error! status: ${response.status}`;
    }

    const data = await response.json();
    
    // Return status for the specified market
    if (market.toLowerCase() === "us") {
      return data.market || "Status unknown";
    } else if (data.exchanges && data.exchanges[market.toLowerCase()]) {
      return data.exchanges[market.toLowerCase()] || "Status unknown";
    } else {
      return "Market not found or status unknown";
    }
  } catch (error) {
    return `Error: ${error.message}`;
  }
}

/**
 * Retrieves company earnings data for a ticker from Polygon.io.
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @param {number} [limit=4] The number of quarters to retrieve.
 * @returns {Promise<string[][]>} A 2D array of earnings data.
 */
export async function getEarnings(ticker, limit = 4) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return "API key not set. Please set your Polygon.io API key.";
    }

    // Check if ticker is valid
    if (!ticker) {
      return "Missing required parameter: ticker";
    }

    limit = (limit && limit > 0) ? limit : 4;
    
    // The current earnings endpoint in Polygon.io's API v3
    const url = `https://api.polygon.io/v3/reference/tickers/${ticker}/results?limit=${limit}&apiKey=${apiKey}`;
    
    const response = await fetch(url);

    if (!response.ok) {
      return [["Earnings data unavailable", `HTTP status: ${response.status}`, "This endpoint might require a premium subscription"]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Header row
      const results = [
        ["Quarter", "Date", "EPS Actual", "EPS Estimate", "Revenue Actual", "Revenue Estimate"]
      ];
      
      // Data rows
      data.results.forEach(quarter => {
        results.push([
          quarter.period || "N/A",
          quarter.date || "N/A",
          quarter.eps || "N/A",
          quarter.eps_estimate || "N/A",
          quarter.revenue || "N/A",
          quarter.revenue_estimate || "N/A"
        ]);
      });
      
      return results;
    } else {
      return [["No earnings data found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Sets the Polygon.io API key for use in all functions.
 * @customfunction
 * @param {string} key The Polygon.io API key.
 * @returns {string} Confirmation message.
 */
export function setPolygonApiKey(key) {
  try {
    if (!key) {
      return "No API key provided.";
    }
    
    apiKey = key;
    localStorage.setItem("polygonApiKey", key);
    return "API key set successfully.";
  } catch (error) {
    return `Error setting API key: ${error.message}`;
  }
}

/**
 * Retrieves dividend information for a ticker from Polygon.io.
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @param {number} [limit=4] The number of dividends to retrieve.
 * @returns {Promise<string[][]>} A 2D array of dividend data.
 */
export async function getDividends(ticker, limit = 4) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return "API key not set. Please set your Polygon.io API key.";
    }

    limit = (limit && limit > 0) ? limit : 4;
    const url = `https://api.polygon.io/v3/reference/dividends?ticker=${ticker}&limit=${limit}&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return `HTTP error! status: ${response.status}`;
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Header row
      const results = [
        ["Ex-Dividend Date", "Payment Date", "Record Date", "Cash Amount", "Declaration Date", "Frequency"]
      ];
      
      // Data rows
      data.results.forEach(dividend => {
        results.push([
          dividend.ex_dividend_date || "N/A",
          dividend.pay_date || "N/A",
          dividend.record_date || "N/A",
          dividend.cash_amount || "N/A",
          dividend.declaration_date || "N/A",
          dividend.frequency || "N/A"
        ]);
      });
      
      return results;
    } else {
      return [["No dividend data found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Calculates technical indicators for a stock.
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @param {string} indicator The indicator to calculate ("SMA", "EMA", "RSI", "MACD").
 * @param {number} [period=14] The period to use for calculation.
 * @param {string} [from] The start date in YYYY-MM-DD format. Defaults to 30 days ago.
 * @param {string} [to] The end date in YYYY-MM-DD format. Defaults to today.
 * @returns {Promise<number[]|string>} An array of indicator values or an error message.
 */
export async function getTechnicalIndicator(ticker, indicator, period = 14, from, to) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return "API key not set. Please set your Polygon.io API key.";
    }

    // Parameter validation
    if (!ticker) {
      return "Missing required parameter: ticker";
    }
    
    if (!indicator) {
      return "Missing required parameter: indicator";
    }

    // Set default dates if not provided
    if (!from) {
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
      from = thirtyDaysAgo.toISOString().split('T')[0];
    }
    
    if (!to) {
      to = new Date().toISOString().split('T')[0];
    }

    // Format validation for dates
    if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) {
      return "Date format must be YYYY-MM-DD";
    }

    // Get historical prices to calculate indicators
    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/day/${from}/${to}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      if (response.status === 403) {
        return "Access forbidden: Check API key permissions or subscription tier";
      } else if (response.status === 429) {
        return "Rate limit exceeded: Try again later or upgrade your subscription";
      }
      return `Could not retrieve price data (HTTP ${response.status})`;
    }

    const data = await response.json();
    if (!data.results || data.results.length === 0) {
      return "Insufficient price data to calculate indicator.";
    }
    
    const prices = data.results.map(bar => bar.c); // Use closing prices
    
    // Validate prices array
    if (!prices.every(price => typeof price === 'number' && !isNaN(price))) {
      return "Invalid price data received";
    }

    // Ensure period is valid
    if (period <= 0 || period >= prices.length) {
      return `Invalid period: ${period}. Must be between 1 and ${prices.length-1}`;
    }

    // Calculate with improved error handling
    try {
      switch (indicator.toUpperCase()) {
        case "SMA":
          return calculateSMA(prices, period);
        case "EMA":
          return calculateEMA(prices, period);
        case "RSI":
          return calculateRSI(prices, period);
        case "MACD":
          return calculateMACD(prices, 12, 26, 9); // Standard MACD parameters
        default:
          return `Unsupported indicator: ${indicator}. Available options: SMA, EMA, RSI, MACD`;
      }
    } catch (calcError) {
      return `Calculation error: ${calcError.message}`;
    }
  } catch (error) {
    return `Error: ${error.message}`;
  }
}

/**
 * Calculates Simple Moving Average.
 * @param {number[]} prices Array of price data.
 * @param {number} period The period to calculate SMA.
 * @returns {number[]} Array of SMA values.
 */
function calculateSMA(prices, period) {
  if (!Array.isArray(prices) || prices.length === 0) {
    throw new Error("Invalid prices array");
  }
  
  if (period <= 0 || period > prices.length) {
    throw new Error(`Invalid period: ${period}`);
  }
  
  const result = [];
  
  // Start from the first complete period
  for (let i = period - 1; i < prices.length; i++) {
    let sum = 0;
    for (let j = 0; j < period; j++) {
      const price = prices[i - j];
      if (typeof price !== 'number' || isNaN(price)) {
        throw new Error(`Invalid price at position ${i-j}: ${price}`);
      }
      sum += price;
    }
    result.push(sum / period);
  }
  
  return result;
}

/**
 * Calculates Exponential Moving Average with improved error handling.
 * @param {number[]} prices Array of price data.
 * @param {number} period The period to calculate EMA.
 * @returns {number[]} Array of EMA values.
 */
function calculateEMA(prices, period) {
  if (!Array.isArray(prices) || prices.length === 0) {
    throw new Error("Invalid prices array");
  }
  
  if (period <= 0 || period > prices.length) {
    throw new Error(`Invalid period: ${period}`);
  }
  
  const result = [];
  const multiplier = 2 / (period + 1);
  
  // Validate all prices first
  for (let i = 0; i < prices.length; i++) {
    if (typeof prices[i] !== 'number' || isNaN(prices[i])) {
      throw new Error(`Invalid price at position ${i}: ${prices[i]}`);
    }
  }
  
  // First EMA is the SMA
  let ema = prices.slice(0, period).reduce((a, b) => a + b, 0) / period;
  result.push(ema);
  
  // Calculate remaining EMAs starting from position 'period'
  for (let i = period; i < prices.length; i++) {
    ema = (prices[i] - ema) * multiplier + ema;
    result.push(ema);
  }
  
  return result;
}

/**
 * Calculates Relative Strength Index with improved error handling.
 * @param {number[]} prices Array of price data.
 * @param {number} period The period to calculate RSI.
 * @returns {number[]} Array of RSI values.
 */
function calculateRSI(prices, period) {
  if (!Array.isArray(prices) || prices.length === 0) {
    throw new Error("Invalid prices array");
  }
  
  if (period <= 0 || period > prices.length - 1) {
    throw new Error(`Invalid period: ${period}`);
  }
  
  const result = [];
  const gains = [];
  const losses = [];
  
  // Calculate price changes
  for (let i = 1; i < prices.length; i++) {
    if (typeof prices[i] !== 'number' || isNaN(prices[i]) || 
        typeof prices[i-1] !== 'number' || isNaN(prices[i-1])) {
      throw new Error(`Invalid price at positions ${i-1} or ${i}`);
    }
    
    const change = prices[i] - prices[i-1];
    gains.push(change > 0 ? change : 0);
    losses.push(change < 0 ? -change : 0);
  }
  
  // Start calculating RSI after we have enough data
  for (let i = period; i < gains.length; i++) {
    // Calculate average gain and loss over the period
    let avgGain = gains.slice(i - period, i).reduce((a, b) => a + b, 0) / period;
    let avgLoss = losses.slice(i - period, i).reduce((a, b) => a + b, 0) / period;
    
    if (avgLoss === 0) {
      result.push(100); // No losses means RSI = 100
    } else {
      const rs = avgGain / avgLoss;
      result.push(100 - (100 / (1 + rs)));
    }
  }
  
  return result;
}

/**
 * Calculates Moving Average Convergence Divergence with improved error handling.
 * @param {number[]} prices Array of price data.
 * @param {number} fastPeriod Fast EMA period.
 * @param {number} slowPeriod Slow EMA period.
 * @param {number} signalPeriod Signal line period.
 * @returns {number[][]} Array containing MACD line, signal line, and histogram.
 */
function calculateMACD(prices, fastPeriod, slowPeriod, signalPeriod) {
  if (!Array.isArray(prices) || prices.length === 0) {
    throw new Error("Invalid prices array");
  }
  
  if (fastPeriod <= 0 || slowPeriod <= 0 || signalPeriod <= 0) {
    throw new Error("Periods must be positive numbers");
  }
  
  if (slowPeriod <= fastPeriod) {
    throw new Error("Slow period must be greater than fast period");
  }
  
  if (prices.length <= slowPeriod) {
    throw new Error(`Not enough price data for the specified periods. Need at least ${slowPeriod + signalPeriod} points.`);
  }
  
  try {
    // Calculate fast and slow EMAs
    const fastEMA = calculateEMA(prices, fastPeriod);
    const slowEMA = calculateEMA(prices, slowPeriod);
    
    // Since the EMAs will have different lengths, we need to align them
    // The slow EMA will be shorter than the fast EMA
    const macdLine = [];
    
    // Calculate MACD line (fast EMA - slow EMA)
    // Skip first (slowPeriod - fastPeriod) elements of fastEMA to align with slowEMA
    const offset = slowPeriod - fastPeriod;
    for (let i = 0; i < slowEMA.length; i++) {
      macdLine.push(fastEMA[i + offset] - slowEMA[i]);
    }
    
    // Calculate signal line (EMA of MACD line)
    // Only use macdLine with enough data points for signalPeriod
    if (macdLine.length < signalPeriod) {
      throw new Error("Insufficient data to calculate signal line");
    }
    
    const signalLine = calculateEMA(macdLine, signalPeriod);
    
    // Calculate histogram (MACD line - signal line)
    // Need to align macdLine with signalLine (signalLine will be shorter)
    const histogram = [];
    const histOffset = macdLine.length - signalLine.length;
    
    for (let i = 0; i < signalLine.length; i++) {
      histogram.push(macdLine[i + histOffset] - signalLine[i]);
    }
    
    // Return arrays of the same length by trimming to the shortest
    const minLength = Math.min(histogram.length, signalLine.length);
    const macdLineAligned = macdLine.slice(macdLine.length - minLength);
    
    return [macdLineAligned, signalLine.slice(0, minLength), histogram.slice(0, minLength)];
  } catch (error) {
    throw new Error(`MACD calculation error: ${error.message}`);
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
