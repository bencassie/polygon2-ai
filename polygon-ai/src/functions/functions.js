/* global clearInterval, console, setInterval */

let apiKey = null;

function getApiKey() {
  return apiKey || localStorage.getItem("polygonApiKey");
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

/****************************************************************************/

/**
 * Searches for tickers matching a search term
 * @customfunction
 * @param {string} searchTerm Text to search for (e.g., "Apple", "Tech", "Bank")
 * @param {number} [limit=10] Maximum number of results to return
 * @returns {Promise<string[][]>} Array of matching tickers and their details
 */
export async function searchTickers(searchTerm, limit = 10) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    limit = Math.min(Math.max(1, limit), 50); // Ensure limit is between 1 and 50
    const url = `https://api.polygon.io/v3/reference/tickers?search=${encodeURIComponent(searchTerm)}&active=true&limit=${limit}&apiKey=${apiKey}`;
    
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Create header row
      const results = [["Symbol", "Name", "Market", "Type", "Primary Exchange"]];
      
      // Data rows
      data.results.forEach(ticker => {
        results.push([
          ticker.ticker || "N/A",
          ticker.name || "N/A",
          ticker.market || "N/A",
          ticker.type || "N/A",
          ticker.primary_exchange || "N/A"
        ]);
      });
      
      return results;
    } else {
      return [["No matching tickers found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves stock splits history for a ticker
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL")
 * @param {number} [limit=10] Maximum number of splits to return
 * @returns {Promise<string[][]>} Array of stock split data
 */
export async function getStockSplits(ticker, limit = 10) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    limit = Math.min(Math.max(1, limit), 50);
    const url = `https://api.polygon.io/v3/reference/splits?ticker=${ticker}&limit=${limit}&apiKey=${apiKey}`;
    
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Create header row
      const results = [["Ticker", "Ex-Date", "Payment Date", "Ratio", "To Factor", "From Factor"]];
      
      // Data rows
      data.results.forEach(split => {
        results.push([
          split.ticker || "N/A",
          split.execution_date || "N/A",
          split.payment_date || "N/A",
          `${split.split_to}:${split.split_from}`,
          split.split_to || "N/A",
          split.split_from || "N/A"
        ]);
      });
      
      return results;
    } else {
      return [["No stock splits found for this ticker."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves historical OHLC data for a specific date range
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL")
 * @param {string} fromDate Start date in YYYY-MM-DD format
 * @param {string} toDate End date in YYYY-MM-DD format
 * @param {string} [timespan="day"] Timespan between data points ("day", "week", "month", "quarter", "year")
 * @returns {Promise<string[][]>} Array of OHLC data
 */
export async function getHistoricalOHLC(ticker, fromDate, toDate, timespan = "day") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Validate dates
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fromDate) || !/^\d{4}-\d{2}-\d{2}$/.test(toDate)) {
      return [["Date format must be YYYY-MM-DD"]];
    }

    // Validate timespan
    const validTimespans = ["day", "week", "month", "quarter", "year"];
    if (!validTimespans.includes(timespan)) {
      return [[`Invalid timespan. Valid options are: ${validTimespans.join(", ")}`]];
    }

    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/${timespan}/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Create header row
      const results = [["Date", "Open", "High", "Low", "Close", "Volume", "VWAP"]];
      
      // Data rows
      data.results.forEach(bar => {
        // Convert timestamp to date
        const date = new Date(bar.t);
        const dateString = date.toISOString().split('T')[0]; // Format as YYYY-MM-DD
        
        results.push([
          dateString,
          bar.o?.toFixed(2) || "N/A", // Open
          bar.h?.toFixed(2) || "N/A", // High
          bar.l?.toFixed(2) || "N/A", // Low
          bar.c?.toFixed(2) || "N/A", // Close
          bar.v || "N/A", // Volume
          bar.vw?.toFixed(2) || "N/A" // VWAP
        ]);
      });
      
      return results;
    } else {
      return [["No data found for this ticker and date range."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves information about stock exchanges
 * @customfunction
 * @param {string} [exchangeType="exchange"] Type of exchange ("exchange", "otc", "index")
 * @returns {Promise<string[][]>} Array of exchange information
 */
export async function getExchanges(exchangeType = "exchange") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Validate exchange type
    const validTypes = ["exchange", "otc", "index"];
    if (!validTypes.includes(exchangeType)) {
      return [[`Invalid exchange type. Valid options are: ${validTypes.join(", ")}`]];
    }

    const url = `https://api.polygon.io/v3/reference/exchanges?asset_class=stocks&apiKey=${apiKey}`;
    
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Filter exchanges by type
      const filteredExchanges = data.results.filter(exchange => 
        exchangeType === "exchange" ? !exchange.type.includes("otc") && !exchange.type.includes("index") :
        exchange.type.includes(exchangeType)
      );

      if (filteredExchanges.length === 0) {
        return [[`No ${exchangeType} exchanges found.`]];
      }

      // Create header row
      const results = [["Name", "Market Identifier Code (MIC)", "Type", "Market", "Country", "Operating MIC"]];
      
      // Data rows
      filteredExchanges.forEach(exchange => {
        results.push([
          exchange.name || "N/A",
          exchange.mic || "N/A",
          exchange.type || "N/A",
          exchange.market || "N/A",
          exchange.locale || "N/A",
          exchange.operating_mic || "N/A"
        ]);
      });
      
      return results;
    } else {
      return [["No exchanges found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}
/**
 * Gets upcoming market holidays and trading hours
 * @customfunction
 * @param {number} [year=current] The year to get holidays for
 * @returns {Promise<string[][]>} Array of market holidays and status
 */
export async function getMarketHolidays(year) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Default to current year if not specified
    if (!year) {
      year = new Date().getFullYear();
    }
    
    const url = `https://api.polygon.io/v1/marketstatus/upcoming?apiKey=${apiKey}`;
    
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const holidays = await response.json();
    if (holidays && holidays.length > 0) {
      // Filter by year if specified
      const filteredHolidays = holidays.filter(holiday => {
        const holidayDate = new Date(holiday.date);
        return holidayDate.getFullYear() === year;
      });

      if (filteredHolidays.length === 0) {
        return [[`No market holidays found for ${year}.`]];
      }

      // Create header row
      const results = [["Date", "Holiday", "Status", "Open", "Close"]];
      
      // Data rows
      filteredHolidays.forEach(holiday => {
        const date = new Date(holiday.date).toISOString().split('T')[0]; // Format as YYYY-MM-DD
        
        results.push([
          date,
          holiday.name || "N/A",
          holiday.status || "N/A",
          holiday.open || "Closed",
          holiday.close || "Closed"
        ]);
      });
      
      return results;
    } else {
      return [["No market holidays found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}
/**
 * Analyzes and compares performance of major market sectors
 * @customfunction
 * @param {string} [timespan="day"] Time period to analyze ("day", "week", "month", "quarter", "year")
 * @returns {Promise<string[][]>} Array of sector performance data
 */
export async function getSectorPerformance(timespan = "day") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Define sector ETFs to track performance
    const sectorETFs = {
      "Technology": "XLK",
      "Financial": "XLF",
      "Healthcare": "XLV",
      "Consumer Discretionary": "XLY",
      "Consumer Staples": "XLP",
      "Energy": "XLE",
      "Materials": "XLB",
      "Industrial": "XLI",
      "Utilities": "XLU",
      "Real Estate": "XLRE",
      "Communication": "XLC"
    };

    // Get S&P 500 for comparison
    const spyTicker = "SPY";

    // Prepare date range based on timespan
    const endDate = new Date();
    const startDate = new Date();
    
    switch(timespan) {
      case "day":
        startDate.setDate(startDate.getDate() - 1);
        break;
      case "week":
        startDate.setDate(startDate.getDate() - 7);
        break;
      case "month":
        startDate.setMonth(startDate.getMonth() - 1);
        break;
      case "quarter":
        startDate.setMonth(startDate.getMonth() - 3);
        break;
      case "year":
        startDate.setFullYear(startDate.getFullYear() - 1);
        break;
      default:
        startDate.setDate(startDate.getDate() - 1);
    }

    const fromDate = startDate.toISOString().split('T')[0];
    const toDate = endDate.toISOString().split('T')[0];

    // Get S&P 500 performance for benchmark
    const spyUrl = `https://api.polygon.io/v2/aggs/ticker/${spyTicker}/range/1/day/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const spyResponse = await fetch(spyUrl);
    
    if (!spyResponse.ok) {
      return [[`HTTP error! status: ${spyResponse.status}`]];
    }
    
    const spyData = await spyResponse.json();
    if (!spyData.results || spyData.results.length < 2) {
      return [["Insufficient S&P 500 data for the selected timespan."]];
    }

    // Calculate S&P 500 performance
    const spyStartPrice = spyData.results[0].c;
    const spyEndPrice = spyData.results[spyData.results.length - 1].c;
    const spyPerformance = ((spyEndPrice / spyStartPrice) - 1) * 100;

    // Prepare results array
    const results = [["Sector", "Performance (%)", "Relative to S&P 500", "Ticker"]];
    
    // Fetch data for each sector ETF
    for (const [sector, ticker] of Object.entries(sectorETFs)) {
      try {
        const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/day/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`;
        const response = await fetch(url);
        
        if (!response.ok) {
          results.push([sector, "Error", "Error", ticker]);
          continue;
        }
        
        const data = await response.json();
        if (!data.results || data.results.length < 2) {
          results.push([sector, "Insufficient data", "N/A", ticker]);
          continue;
        }
        
        const startPrice = data.results[0].c;
        const endPrice = data.results[data.results.length - 1].c;
        const performance = ((endPrice / startPrice) - 1) * 100;
        const relativePerformance = performance - spyPerformance;
        
        results.push([
          sector,
          performance.toFixed(2) + "%",
          (relativePerformance > 0 ? "+" : "") + relativePerformance.toFixed(2) + "%",
          ticker
        ]);
      } catch (error) {
        results.push([sector, `Error: ${error.message}`, "N/A", ticker]);
      }
    }
    
    // Sort by performance (descending)
    results.sort((a, b) => {
      // Skip header row
      if (a[0] === "Sector") return -1;
      if (b[0] === "Sector") return 1;
      
      // Parse percentages for sorting
      const aPerf = parseFloat(a[1]);
      const bPerf = parseFloat(b[1]);
      
      return isNaN(aPerf) || isNaN(bPerf) ? 0 : bPerf - aPerf;
    });
    
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Calculates correlation between two stocks over a period
 * @customfunction
 * @param {string} ticker1 First ticker symbol
 * @param {string} ticker2 Second ticker symbol
 * @param {number} [days=30] Number of trading days to analyze
 * @returns {Promise<number>} Correlation coefficient (-1 to 1)
 */
export async function getStockCorrelation(ticker1, ticker2, days = 30) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return "API key not set. Please set your Polygon.io API key.";
    }

    // Limit days to a reasonable range
    days = Math.min(Math.max(5, days), 365);

    // Calculate date range
    const endDate = new Date();
    const startDate = new Date();
    // Add buffer days to ensure we get enough trading days
    startDate.setDate(startDate.getDate() - (days * 2)); 
    
    const fromDate = startDate.toISOString().split('T')[0];
    const toDate = endDate.toISOString().split('T')[0];

    // Fetch data for both tickers
    const responses = await Promise.all([
      fetch(`https://api.polygon.io/v2/aggs/ticker/${ticker1}/range/1/day/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`),
      fetch(`https://api.polygon.io/v2/aggs/ticker/${ticker2}/range/1/day/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`)
    ]);

    if (!responses[0].ok || !responses[1].ok) {
      return `HTTP error! status: ${responses[0].status || responses[1].status}`;
    }

    const data = await Promise.all([
      responses[0].json(),
      responses[1].json()
    ]);

    if (!data[0].results || !data[1].results || 
        data[0].results.length < days || data[1].results.length < days) {
      return "Insufficient price data for correlation calculation.";
    }

    // Take the most recent 'days' days of data
    const prices1 = data[0].results.slice(-days).map(bar => bar.c);
    const prices2 = data[1].results.slice(-days).map(bar => bar.c);

    // Calculate daily returns
    const returns1 = [];
    const returns2 = [];
    
    for (let i = 1; i < prices1.length; i++) {
      returns1.push((prices1[i] / prices1[i-1]) - 1);
      returns2.push((prices2[i] / prices2[i-1]) - 1);
    }

    // Calculate correlation coefficient
    const correlation = calculateCorrelation(returns1, returns2);
    return parseFloat(correlation.toFixed(4)); // Return as number with 4 decimal places
  } catch (error) {
    return `Error: ${error.message}`;
  }
}

/**
 * Calculates Pearson correlation coefficient between two arrays
 * @private
 * @param {number[]} x First array
 * @param {number[]} y Second array
 * @returns {number} Correlation coefficient
 */
function calculateCorrelation(x, y) {
  const n = x.length;
  
  // Calculate means
  const xMean = x.reduce((a, b) => a + b, 0) / n;
  const yMean = y.reduce((a, b) => a + b, 0) / n;
  
  // Calculate sums for correlation formula
  let numerator = 0;
  let xDenom = 0;
  let yDenom = 0;
  
  for (let i = 0; i < n; i++) {
    const xDiff = x[i] - xMean;
    const yDiff = y[i] - yMean;
    numerator += xDiff * yDiff;
    xDenom += xDiff * xDiff;
    yDenom += yDiff * yDiff;
  }
  
  const denominator = Math.sqrt(xDenom * yDenom);
  
  return denominator === 0 ? 0 : numerator / denominator;
}
/**
 * Creates a simple portfolio tracker with current values and returns
 * @customfunction
 * @param {string[][]} portfolioData 2D array of [ticker, shares]
 * @returns {Promise<string[][]>} Portfolio summary with current value and performance
 */
export async function getPortfolioSummary(portfolioData) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    if (!Array.isArray(portfolioData) || portfolioData.length === 0) {
      return [["Invalid portfolio data. Expected format: [[ticker, shares], ...]"]];
    }

    // Create header row
    const results = [["Symbol", "Shares", "Current Price", "Market Value", "Day Change %"]];
    let totalValue = 0;
    
    // Process each position
    for (const position of portfolioData) {
      if (!Array.isArray(position) || position.length < 2) {
        results.push(["Invalid position", "N/A", "N/A", "N/A", "N/A"]);
        continue;
      }
      
      const [ticker, sharesInput] = position;
      const shares = parseFloat(sharesInput);
      
      if (isNaN(shares)) {
        results.push([ticker, "Invalid", "N/A", "N/A", "N/A"]);
        continue;
      }
      
      try {
        const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/prev?adjusted=true&apiKey=${apiKey}`;
        const response = await fetch(url);
        
        if (!response.ok) {
          results.push([ticker, shares, "Error", "Error", "Error"]);
          continue;
        }
        
        const data = await response.json();
        if (!data.results || data.results.length === 0) {
          results.push([ticker, shares, "No data", "N/A", "N/A"]);
          continue;
        }
        
        const priceData = data.results[0];
        const currentPrice = priceData.c;
        const previousClose = priceData.o; // Using open price as previous reference
        const marketValue = currentPrice * shares;
        const dayChange = ((currentPrice / previousClose) - 1) * 100;
        
        totalValue += marketValue;
        
        results.push([
          ticker,
          shares.toString(),
          "$" + currentPrice.toFixed(2),
          "$" + marketValue.toFixed(2),
          (dayChange > 0 ? "+" : "") + dayChange.toFixed(2) + "%"
        ]);
      } catch (error) {
        results.push([ticker, shares, `Error: ${error.message}`, "N/A", "N/A"]);
      }
    }
    
    // Add total row
    results.push(["Total", "", "", "$" + totalValue.toFixed(2), ""]);
    
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
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
 * @returns {string[][]} A 2D array of indicator values with dates or an error message.
 */
export async function getTechnicalIndicator(ticker, indicator, period = 14, from, to) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Parameter validation
    if (!ticker) {
      return [["Missing required parameter: ticker"]];
    }
    
    if (!indicator) {
      return [["Missing required parameter: indicator"]];
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
      return [["Date format must be YYYY-MM-DD"]];
    }

    // Get historical prices to calculate indicators
    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/day/${from}/${to}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      if (response.status === 403) {
        return [["Access forbidden: Check API key permissions or subscription tier"]];
      } else if (response.status === 429) {
        return [["Rate limit exceeded: Try again later or upgrade your subscription"]];
      }
      return [[`Could not retrieve price data (HTTP ${response.status})`]];
    }

    const data = await response.json();
    if (!data.results || data.results.length === 0) {
      return [["Insufficient price data to calculate indicator."]];
    }
    
    const prices = data.results.map(bar => bar.c); // Use closing prices
    const dates = data.results.map(bar => {
      // Convert timestamp to date string (timestamp is in milliseconds)
      const date = new Date(bar.t);
      return date.toISOString().split('T')[0]; // Format as YYYY-MM-DD
    });
    
    // Validate prices array
    if (!prices.every(price => typeof price === 'number' && !isNaN(price))) {
      return [["Invalid price data received"]];
    }

    // Ensure period is valid
    if (period <= 0 || period >= prices.length) {
      return [[`Invalid period: ${period}. Must be between 1 and ${prices.length-1}`]];
    }

    // Calculate with improved error handling
    try {
      let result = [];
      let indicatorValues = [];
      let startIndex = 0;
      
      switch (indicator.toUpperCase()) {
        case "SMA":
          indicatorValues = calculateSMA(prices, period);
          startIndex = period - 1;
          // Create a header row
          result.push(["Date", `SMA(${period})`]);
          break;
        case "EMA":
          indicatorValues = calculateEMA(prices, period);
          startIndex = period - 1;
          result.push(["Date", `EMA(${period})`]);
          break;
        case "RSI":
          indicatorValues = calculateRSI(prices, period);
          startIndex = period;
          result.push(["Date", `RSI(${period})`]);
          break;
        case "MACD":
          const macdResult = calculateMACD(prices, 12, 26, 9); // Standard MACD parameters
          const [macdLine, signalLine, histogram] = macdResult;
          
          // For MACD, we need to determine the start index based on the shortest array
          startIndex = prices.length - macdLine.length;
          
          // Create header row for MACD
          result.push(["Date", "MACD Line", "Signal Line", "Histogram"]);
          
          // Create data rows for MACD (all 3 components)
          for (let i = 0; i < macdLine.length; i++) {
            const dateIndex = startIndex + i;
            if (dateIndex < dates.length) {
              result.push([
                dates[dateIndex],
                macdLine[i].toFixed(4),
                signalLine[i].toFixed(4),
                histogram[i].toFixed(4)
              ]);
            }
          }
          
          return result;
        default:
          return [[`Unsupported indicator: ${indicator}. Available options: SMA, EMA, RSI, MACD`]];
      }
      
      // For single-value indicators (SMA, EMA, RSI), create rows with date and value
      for (let i = 0; i < indicatorValues.length; i++) {
        const dateIndex = startIndex + i;
        if (dateIndex < dates.length) {
          result.push([
            dates[dateIndex],
            indicatorValues[i].toFixed(4)
          ]);
        }
      }
      
      return result;
    } catch (calcError) {
      return [[`Calculation error: ${calcError.message}`]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
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
