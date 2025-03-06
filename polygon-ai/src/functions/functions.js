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

/**
 * Retrieves batch quotes for multiple tickers
 * @customfunction
 * @param {string[]} tickers Array of stock ticker symbols
 * @returns {Promise<string[][]>} Array of latest quotes for each ticker
 */
export async function getBatchQuotes(tickers) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    const results = [["Ticker", "Price", "Change %", "Volume", "Last Updated"]];
    
    // Process tickers in batches of 5 to respect rate limits
    for (let i = 0; i < tickers.length; i += 5) {
      const batch = tickers.slice(i, i + 5);
      
      // Get quotes for current batch
      const promises = batch.map(ticker =>
        fetch(`https://api.polygon.io/v2/last/trade/${ticker}?apiKey=${apiKey}`)
          .then(response => response.json())
      );

      const responses = await Promise.all(promises);

      responses.forEach((data, index) => {
        if (data.results && data.results.length > 0) {
          const quote = data.results[0];
          const changePercent = ((quote.c / quote.o) - 1) * 100;
          const lastUpdated = new Date(quote.t).toLocaleString();
          
          results.push([
            batch[index],
            quote.p.toFixed(2),
            changePercent.toFixed(2) + "%",
            quote.s.toLocaleString(),
            lastUpdated
          ]);
        } else {
          results.push([batch[index], "No data", "N/A", "N/A", "N/A"]);
        }
      });

      // Add delay between batches to respect rate limits
      if (i + 5 < tickers.length) {
        await new Promise(resolve => setTimeout(resolve, 200));
      }
    }
    
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves options chain data for a stock
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @param {string} expirationDate Expiration date in YYYY-MM-DD format
 * @returns {Promise<string[][]>} Array of options data
 */
export async function getOptionsChain(ticker, expirationDate) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    if (!/^\d{4}-\d{2}-\d{2}$/.test(expirationDate)) {
      return [["Invalid date format. Use YYYY-MM-DD"]];
    }

    const url = `https://api.polygon.io/v3/reference/options/contracts?underlying_ticker=${ticker}&expiration_date=${expirationDate}&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Create header row
      const results = [["Contract", "Type", "Strike", "Last Price", "Volume", "Open Interest"]];
      
      // Sort options by strike price and type
      const sortedOptions = data.results.sort((a, b) => {
        if (a.strike_price === b.strike_price) {
          return a.type.localeCompare(b.type);
        }
        return a.strike_price - b.strike_price;
      });

      // Add data rows
      for (const option of sortedOptions) {
        results.push([
          option.ticker,
          option.type.toUpperCase(),
          "$" + option.strike_price.toFixed(2),
          option.last_price ? "$" + option.last_price.toFixed(2) : "N/A",
          option.volume?.toLocaleString() || "N/A",
          option.open_interest?.toLocaleString() || "N/A"
        ]);
      }

      return results;
    } else {
      return [["No options data found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Calculates options Greeks using Black-Scholes model
 * @customfunction
 * @param {string} optionTicker The option contract ticker
 * @returns {Promise<string[][]>} Array of options Greeks data
 */
export async function getOptionsGreeks(optionTicker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    const url = `https://api.polygon.io/v3/reference/options/contracts/${optionTicker}?apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results) {
      const option = data.results;
      
      // Create header row
      const results = [["Greek", "Value", "Description"]];
      
      // Add Greeks data
      results.push(
        ["Delta", option.delta?.toFixed(4) || "N/A", "Option's sensitivity to underlying price change"],
        ["Gamma", option.gamma?.toFixed(4) || "N/A", "Rate of change in delta"],
        ["Theta", option.theta?.toFixed(4) || "N/A", "Time decay of option value"],
        ["Vega", option.vega?.toFixed(4) || "N/A", "Sensitivity to volatility changes"],
        ["Rho", option.rho?.toFixed(4) || "N/A", "Sensitivity to interest rate changes"],
        ["IV", (option.implied_volatility * 100)?.toFixed(2) + "%" || "N/A", "Implied Volatility"]
      );

      return results;
    } else {
      return [["No options data found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves forex exchange rates
 * @customfunction
 * @param {string} fromCurrency Base currency code (e.g., "USD")
 * @param {string} toCurrency Quote currency code (e.g., "EUR")
 * @returns {Promise<string[][]>} Array of forex rate data
 */
export async function getForexRates(fromCurrency, toCurrency) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    const url = `https://api.polygon.io/v2/aggs/ticker/C:${fromCurrency}${toCurrency}/prev?adjusted=true&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      const rate = data.results[0];
      
      // Create results array with rate information
      return [
        ["Currency Pair", "Rate", "Daily Change", "Last Updated"],
        [
          `${fromCurrency}/${toCurrency}`,
          rate.c.toFixed(4),
          ((rate.c / rate.o - 1) * 100).toFixed(2) + "%",
          new Date(rate.t).toLocaleString()
        ]
      ];
    } else {
      return [["No forex data found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves crypto market data
 * @customfunction
 * @param {string} cryptoPair Trading pair (e.g., "BTC/USD")
 * @returns {Promise<string[][]>} Array of crypto market data
 */
export async function getCryptoData(cryptoPair) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    const formattedPair = cryptoPair.replace("/", "-");
    const url = `https://api.polygon.io/v2/aggs/ticker/X:${formattedPair}/prev?adjusted=true&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      const crypto = data.results[0];
      
      return [
        ["Trading Pair", "Price", "24h High", "24h Low", "24h Volume", "24h Change"],
        [
          cryptoPair,
          "$" + crypto.c.toFixed(2),
          "$" + crypto.h.toFixed(2),
          "$" + crypto.l.toFixed(2),
          crypto.v.toLocaleString(),
          ((crypto.c / crypto.o - 1) * 100).toFixed(2) + "%"
        ]
      ];
    } else {
      return [["No crypto data found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves market depth data (Level 2 book)
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @returns {Promise<string[][]>} Array of market depth data
 */
export async function getMarketDepth(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    const url = `https://api.polygon.io/v3/quotes/${ticker}?limit=100&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Create header rows
      const results = [
        ["Bid Price", "Bid Size", "", "Ask Price", "Ask Size"],
        ["=========", "========", "", "=========", "========"]
      ];
      
      // Group quotes by bid and ask
      const bids = new Map();
      const asks = new Map();
      
      data.results.forEach(quote => {
        if (quote.bid_price > 0) {
          bids.set(quote.bid_price, (bids.get(quote.bid_price) || 0) + quote.bid_size);
        }
        if (quote.ask_price > 0) {
          asks.set(quote.ask_price, (asks.get(quote.ask_price) || 0) + quote.ask_size);
        }
      });

      // Sort bids (descending) and asks (ascending)
      const sortedBids = [...bids.entries()].sort((a, b) => b[0] - a[0]);
      const sortedAsks = [...asks.entries()].sort((a, b) => a[0] - b[0]);
      
      // Combine bid and ask data
      const maxRows = Math.max(sortedBids.length, sortedAsks.length);
      for (let i = 0; i < maxRows; i++) {
        const bid = sortedBids[i] || ["", ""];
        const ask = sortedAsks[i] || ["", ""];
        
        results.push([
          bid[0]?.toFixed(2) || "",
          bid[1]?.toLocaleString() || "",
          "",
          ask[0]?.toFixed(2) || "",
          ask[1]?.toLocaleString() || ""
        ]);
      }

      return results;
    } else {
      return [["No market depth data found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves market breadth indicators
 * @customfunction
 * @param {string} [market="stocks"] Market to analyze ("stocks", "crypto", "forex")
 * @returns {Promise<string[][]>} Array of market breadth indicators
 */
export async function getMarketBreadth(market = "stocks") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Get snapshot of all tickers in the market
    const url = `https://api.polygon.io/v2/snapshot/locale/us/markets/stocks/tickers?apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.tickers && data.tickers.length > 0) {
      let advancing = 0;
      let declining = 0;
      let unchanged = 0;
      let newHighs = 0;
      let newLows = 0;
      let totalVolume = 0;

      data.tickers.forEach(ticker => {
        const day = ticker.day;
        if (day) {
          if (day.c > day.o) advancing++;
          else if (day.c < day.o) declining++;
          else unchanged++;

          if (day.h >= ticker.prevDay?.h) newHighs++;
          if (day.l <= ticker.prevDay?.l) newLows++;
          totalVolume += day.v || 0;
        }
      });

      const advDecRatio = (advancing / (declining || 1)).toFixed(2);
      const breadthIndex = ((advancing - declining) / (advancing + declining) * 100).toFixed(2);

      return [
        ["Indicator", "Value"],
        ["Advancing Issues", advancing.toString()],
        ["Declining Issues", declining.toString()],
        ["Unchanged Issues", unchanged.toString()],
        ["New 52-Week Highs", newHighs.toString()],
        ["New 52-Week Lows", newLows.toString()],
        ["Advance/Decline Ratio", advDecRatio],
        ["Market Breadth Index", breadthIndex + "%"],
        ["Total Volume", totalVolume.toLocaleString()]
      ];
    } else {
      return [["No market breadth data found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves corporate actions (dividends, splits, etc.)
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @returns {Promise<string[][]>} Array of corporate actions data
 */
export async function getCorporateActions(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    const url = `https://api.polygon.io/v3/reference/splits?ticker=${ticker}&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Create header row
      const results = [["Date", "Type", "Description", "Factor"]];
      
      // Add data rows
      data.results.forEach(action => {
        results.push([
          action.execution_date,
          "Stock Split",
          `${action.split_to}:${action.split_from} Split`,
          (action.split_to / action.split_from).toFixed(2)
        ]);
      });

      return results;
    } else {
      return [["No corporate actions found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves analyst recommendations
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @returns {Promise<string[][]>} Array of analyst recommendations
 */
export async function getAnalystRecommendations(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Get price targets using the correct v3 endpoint
    const url = `https://api.polygon.io/v3/reference/analysts/price-targets/${ticker}?apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    
    // Create header row
    const results = [["Date", "Analyst", "Price Target", "Previous Target", "Change"]];

    if (data.results && data.results.length > 0) {
      // Sort recommendations by date (newest first)
      const sortedRecs = data.results.sort((a, b) => 
        new Date(b.published_date) - new Date(a.published_date)
      );

      // Add data rows
      sortedRecs.forEach(rec => {
        const change = rec.price_target - (rec.prior_price_target || rec.price_target);
        const changeStr = change === 0 ? "No Change" : 
          (change > 0 ? "+" + change.toFixed(2) : change.toFixed(2));

        results.push([
          rec.published_date.split('T')[0],
          rec.analyst_name || rec.publisher_name || "Unknown",
          "$" + rec.price_target.toFixed(2),
          rec.prior_price_target ? "$" + rec.prior_price_target.toFixed(2) : "N/A",
          changeStr
        ]);
      });

      return results;
    } else {
      return [["No analyst recommendations found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves earnings analysis and estimates
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @returns {Promise<string[][]>} Array of earnings data
 */
export async function getEarningsAnalysis(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    const url = `https://api.polygon.io/v3/reference/earnings?ticker=${ticker}&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Create header rows
      const results = [
        ["Fiscal Quarter", "Report Date", "EPS Est.", "EPS Actual", "Revenue Est.", "Revenue Actual", "Surprise %"]
      ];
      
      // Sort earnings by date (newest first)
      const sortedEarnings = data.results.sort((a, b) => 
        new Date(b.period_end_date) - new Date(a.period_end_date)
      );

      // Add data rows
      sortedEarnings.forEach(earning => {
        const epsSurprise = earning.eps_actual && earning.eps_estimate
          ? ((earning.eps_actual - earning.eps_estimate) / Math.abs(earning.eps_estimate) * 100).toFixed(2)
          : "N/A";

        results.push([
          `${earning.fiscal_year}Q${earning.fiscal_quarter}`,
          earning.period_end_date,
          earning.eps_estimate?.toFixed(2) || "N/A",
          earning.eps_actual?.toFixed(2) || "N/A",
          earning.revenue_estimate ? "$" + (earning.revenue_estimate / 1e6).toFixed(2) + "M" : "N/A",
          earning.revenue_actual ? "$" + (earning.revenue_actual / 1e6).toFixed(2) + "M" : "N/A",
          epsSurprise !== "N/A" ? epsSurprise + "%" : "N/A"
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
 * Calculates Bollinger Bands for a stock
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @param {number} [period=20] Period for moving average
 * @param {number} [stdDev=2] Number of standard deviations
 * @returns {Promise<string[][]>} Array of Bollinger Bands data
 */
export async function getBollingerBands(ticker, period = 20, stdDev = 2) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Get historical prices
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - (period * 2)); // Double the period for better calculation

    const fromDate = startDate.toISOString().split('T')[0];
    const toDate = endDate.toISOString().split('T')[0];

    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/day/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length >= period) {
      // Create header row
      const results = [["Date", "Price", "Middle Band", "Upper Band", "Lower Band"]];
      
      // Calculate bands for each period
      for (let i = period - 1; i < data.results.length; i++) {
        const prices = data.results.slice(i - period + 1, i + 1).map(bar => bar.c);
        
        // Calculate SMA
        const sma = prices.reduce((a, b) => a + b, 0) / period;
        
        // Calculate standard deviation
        const variance = prices.reduce((a, b) => a + Math.pow(b - sma, 2), 0) / period;
        const standardDev = Math.sqrt(variance);
        
        // Calculate bands
        const upperBand = sma + (standardDev * stdDev);
        const lowerBand = sma - (standardDev * stdDev);
        
        // Format date
        const date = new Date(data.results[i].t);
        const dateStr = date.toISOString().split('T')[0];
        
        results.push([
          dateStr,
          data.results[i].c.toFixed(2),
          sma.toFixed(2),
          upperBand.toFixed(2),
          lowerBand.toFixed(2)
        ]);
      }
      
      return results;
    } else {
      return [["Insufficient data for Bollinger Bands calculation."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Calculates Average True Range (ATR)
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @param {number} [period=14] Period for ATR calculation
 * @returns {Promise<string[][]>} Array of ATR data
 */
export async function getATR(ticker, period = 14) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Get historical prices
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - (period * 2)); // Double the period for better calculation

    const fromDate = startDate.toISOString().split('T')[0];
    const toDate = endDate.toISOString().split('T')[0];

    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/day/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length >= period) {
      // Create header row
      const results = [["Date", "ATR", "True Range"]];
      
      let atr = 0;
      let prevClose = data.results[0].c;
      
      // Calculate initial ATR
      const trueRanges = [];
      for (let i = 1; i < period; i++) {
        const high = data.results[i].h;
        const low = data.results[i].l;
        const close = data.results[i].c;
        
        // Calculate True Range
        const tr = Math.max(
          high - low,
          Math.abs(high - prevClose),
          Math.abs(low - prevClose)
        );
        
        trueRanges.push(tr);
        prevClose = close;
      }
      
      // Initial ATR is simple average of first 'period' true ranges
      atr = trueRanges.reduce((a, b) => a + b, 0) / period;
      
      // Calculate ATR for remaining periods
      for (let i = period; i < data.results.length; i++) {
        const high = data.results[i].h;
        const low = data.results[i].l;
        const close = data.results[i].c;
        
        // Calculate True Range
        const tr = Math.max(
          high - low,
          Math.abs(high - prevClose),
          Math.abs(low - prevClose)
        );
        
        // Calculate ATR using smoothing formula
        atr = ((atr * (period - 1)) + tr) / period;
        
        // Format date
        const date = new Date(data.results[i].t);
        const dateStr = date.toISOString().split('T')[0];
        
        results.push([
          dateStr,
          atr.toFixed(2),
          tr.toFixed(2)
        ]);
        
        prevClose = close;
      }
      
      return results;
    } else {
      return [["Insufficient data for ATR calculation."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Calculates Fibonacci retracement levels
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @param {string} fromDate Start date in YYYY-MM-DD format
 * @param {string} toDate End date in YYYY-MM-DD format
 * @returns {Promise<string[][]>} Array of Fibonacci levels
 */
export async function getFibonacciLevels(ticker, fromDate, toDate) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Validate dates
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fromDate) || !/^\d{4}-\d{2}-\d{2}$/.test(toDate)) {
      return [["Date format must be YYYY-MM-DD"]];
    }

    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/day/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Find high and low in the period
      const prices = data.results.map(bar => ({high: bar.h, low: bar.l}));
      const high = Math.max(...prices.map(p => p.high));
      const low = Math.min(...prices.map(p => p.low));
      const diff = high - low;

      // Fibonacci ratios
      const levels = [
        {ratio: 0, level: high},
        {ratio: 0.236, level: high - (diff * 0.236)},
        {ratio: 0.382, level: high - (diff * 0.382)},
        {ratio: 0.5, level: high - (diff * 0.5)},
        {ratio: 0.618, level: high - (diff * 0.618)},
        {ratio: 0.786, level: high - (diff * 0.786)},
        {ratio: 1, level: low}
      ];

      // Create results array
      const results = [
        ["Fibonacci Ratio", "Price Level", "Description"]
      ];

      // Add levels with descriptions
      levels.forEach(({ratio, level}) => {
        results.push([
          ratio.toFixed(3),
          "$" + level.toFixed(2),
          getFibonacciDescription(ratio)
        ]);
      });

      // Add period high/low info
      results.push(
        ["", "", ""],
        ["Period High", "$" + high.toFixed(2), fromDate + " to " + toDate],
        ["Period Low", "$" + low.toFixed(2), ""]
      );

      return results;
    } else {
      return [["No price data found for the specified period."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Helper function to get Fibonacci level descriptions
 * @private
 */
function getFibonacciDescription(ratio) {
  switch (ratio) {
    case 0: return "Swing High (0%)";
    case 0.236: return "Weak Retracement (23.6%)";
    case 0.382: return "Initial Support/Resistance (38.2%)";
    case 0.5: return "Middle Retracement (50%)";
    case 0.618: return "Golden Ratio (61.8%)";
    case 0.786: return "Deep Retracement (78.6%)";
    case 1: return "Swing Low (100%)";
    default: return "";
  }
}

/**
 * Analyzes options strategies and calculates potential returns
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @param {string} strategy Strategy type ("covered_call", "iron_condor", "butterfly")
 * @returns {Promise<string[][]>} Array of strategy analysis data
 */
export async function getOptionsStrategy(ticker, strategy) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Get current stock price
    const priceUrl = `https://api.polygon.io/v2/aggs/ticker/${ticker}/prev?adjusted=true&apiKey=${apiKey}`;
    const priceResponse = await fetch(priceUrl);
    
    if (!priceResponse.ok) {
      return [[`HTTP error! status: ${priceResponse.status}`]];
    }

    const priceData = await priceResponse.json();
    if (!priceData.results || !priceData.results[0]) {
      return [["No price data found."]];
    }

    const currentPrice = priceData.results[0].c;
    const results = [["Strategy Component", "Strike", "Type", "Premium", "Max Profit", "Max Loss"]];

    // Get nearest expiration options chain
    const today = new Date();
    const thirtyDaysFromNow = new Date();
    thirtyDaysFromNow.setDate(today.getDate() + 30);
    
    const optionsUrl = `https://api.polygon.io/v3/reference/options/contracts?underlying_ticker=${ticker}&expiration_date.lte=${thirtyDaysFromNow.toISOString().split('T')[0]}&expiration_date.gte=${today.toISOString().split('T')[0]}&apiKey=${apiKey}`;
    const optionsResponse = await fetch(optionsUrl);

    if (!optionsResponse.ok) {
      return [[`HTTP error! status: ${optionsResponse.status}`]];
    }

    const optionsData = await optionsResponse.json();
    if (!optionsData.results || optionsData.results.length === 0) {
      return [["No options data found."]];
    }

    // Sort options by strike price
    const options = optionsData.results.sort((a, b) => a.strike_price - b.strike_price);
    
    // Find ATM options (closest to current price)
    const atmIndex = options.findIndex(opt => opt.strike_price >= currentPrice);
    
    switch (strategy.toLowerCase()) {
      case "covered_call": {
        const atmCall = options[atmIndex];
        if (!atmCall) return [["Could not find appropriate options for strategy."]];
        
        results.push(
          ["Stock Purchase", "$" + currentPrice.toFixed(2), "LONG", "N/A", "", ""],
          ["Call Sale", "$" + atmCall.strike_price.toFixed(2), "SHORT", "$" + (atmCall.last_price || 0).toFixed(2), "", ""],
          ["", "", "", "", "", ""],
          ["Total", "", "", "", 
           "$" + ((atmCall.strike_price - currentPrice + (atmCall.last_price || 0)) * 100).toFixed(2),
           "$" + ((currentPrice - (atmCall.last_price || 0)) * -100).toFixed(2)]
        );
        break;
      }
      
      case "iron_condor": {
        const lowerPut = options[Math.max(0, atmIndex - 2)];
        const upperPut = options[Math.max(0, atmIndex - 1)];
        const lowerCall = options[Math.min(options.length - 1, atmIndex + 1)];
        const upperCall = options[Math.min(options.length - 1, atmIndex + 2)];
        
        if (!lowerPut || !upperPut || !lowerCall || !upperCall) {
          return [["Could not find appropriate options for strategy."]];
        }
        
        const maxCredit = ((upperPut.last_price || 0) - (lowerPut.last_price || 0) + 
                          (lowerCall.last_price || 0) - (upperCall.last_price || 0)).toFixed(2);
        const maxLoss = ((upperPut.strike_price - lowerPut.strike_price) * 100 - maxCredit * 100).toFixed(2);
        
        results.push(
          ["Put Spread Lower", "$" + lowerPut.strike_price.toFixed(2), "LONG", "$" + (lowerPut.last_price || 0).toFixed(2), "", ""],
          ["Put Spread Upper", "$" + upperPut.strike_price.toFixed(2), "SHORT", "$" + (upperPut.last_price || 0).toFixed(2), "", ""],
          ["Call Spread Lower", "$" + lowerCall.strike_price.toFixed(2), "SHORT", "$" + (lowerCall.last_price || 0).toFixed(2), "", ""],
          ["Call Spread Upper", "$" + upperCall.strike_price.toFixed(2), "LONG", "$" + (upperCall.last_price || 0).toFixed(2), "", ""],
          ["", "", "", "", "", ""],
          ["Total", "", "", "$" + maxCredit, "$" + (maxCredit * 100).toFixed(2), "$" + maxLoss]
        );
        break;
      }
      
      case "butterfly": {
        const lowerStrike = options[Math.max(0, atmIndex - 1)];
        const middleStrike = options[atmIndex];
        const upperStrike = options[Math.min(options.length - 1, atmIndex + 1)];
        
        if (!lowerStrike || !middleStrike || !upperStrike) {
          return [["Could not find appropriate options for strategy."]];
        }
        
        const maxRisk = ((lowerStrike.last_price || 0) - 2 * (middleStrike.last_price || 0) + (upperStrike.last_price || 0)) * 100;
        const maxReward = (middleStrike.strike_price - lowerStrike.strike_price) * 100 - maxRisk;
        
        results.push(
          ["Lower Wing", "$" + lowerStrike.strike_price.toFixed(2), "LONG", "$" + (lowerStrike.last_price || 0).toFixed(2), "", ""],
          ["Body", "$" + middleStrike.strike_price.toFixed(2), "SHORT x2", "$" + (middleStrike.last_price || 0).toFixed(2), "", ""],
          ["Upper Wing", "$" + upperStrike.strike_price.toFixed(2), "LONG", "$" + (upperStrike.last_price || 0).toFixed(2), "", ""],
          ["", "", "", "", "", ""],
          ["Total", "", "", "", "$" + maxReward.toFixed(2), "$" + maxRisk.toFixed(2)]
        );
        break;
      }
      
      default:
        return [[`Unsupported strategy: ${strategy}. Available strategies: covered_call, iron_condor, butterfly`]];
    }

    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Calculates various risk indicators for a stock
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @returns {Promise<string[][]>} Array of risk metrics
 */
export async function getRiskIndicators(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Get historical prices for the past year
    const endDate = new Date();
    const startDate = new Date();
    startDate.setFullYear(startDate.getFullYear() - 1);

    const fromDate = startDate.toISOString().split('T')[0];
    const toDate = endDate.toISOString().split('T')[0];

    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/day/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Calculate daily returns
      const returns = [];
      for (let i = 1; i < data.results.length; i++) {
        returns.push((data.results[i].c / data.results[i-1].c) - 1);
      }

      // Calculate risk metrics
      const annualReturn = (Math.pow(1 + returns.reduce((a, b) => a + b, 0) / returns.length, 252) - 1) * 100;
      const dailyVolatility = Math.sqrt(returns.reduce((a, b) => a + Math.pow(b - (returns.reduce((x, y) => x + y, 0) / returns.length), 2), 0) / returns.length);
      const annualVolatility = dailyVolatility * Math.sqrt(252) * 100;
      
      // Calculate drawdown
      let maxDrawdown = 0;
      let peak = data.results[0].c;
      
      data.results.forEach(bar => {
        if (bar.c > peak) {
          peak = bar.c;
        }
        const drawdown = (peak - bar.c) / peak;
        maxDrawdown = Math.max(maxDrawdown, drawdown);
      });

      // Calculate Sharpe Ratio (assuming risk-free rate of 2%)
      const riskFreeRate = 0.02;
      const excessReturn = annualReturn / 100 - riskFreeRate;
      const sharpeRatio = excessReturn / (annualVolatility / 100);

      // Calculate Value at Risk (95% confidence)
      const sortedReturns = returns.sort((a, b) => a - b);
      const varIndex = Math.floor(returns.length * 0.05);
      const valueAtRisk = -sortedReturns[varIndex] * 100;

      return [
        ["Risk Metric", "Value", "Description"],
        ["Annual Return", annualReturn.toFixed(2) + "%", "Annualized return over the period"],
        ["Annual Volatility", annualVolatility.toFixed(2) + "%", "Annualized standard deviation of returns"],
        ["Maximum Drawdown", (maxDrawdown * 100).toFixed(2) + "%", "Largest peak-to-trough decline"],
        ["Sharpe Ratio", sharpeRatio.toFixed(2), "Risk-adjusted return measure"],
        ["Value at Risk (95%)", valueAtRisk.toFixed(2) + "%", "Potential loss at 95% confidence level"],
        ["Daily VaR ($)", "$" + (data.results[data.results.length-1].c * valueAtRisk / 100).toFixed(2), "Dollar value at risk per $100 invested"],
        ["Beta", "", "Calculated separately using market comparison"],
        ["Alpha", "", "Calculated separately using risk-free rate and market returns"]
      ];
    } else {
      return [["Insufficient data for risk analysis."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves social sentiment metrics for a stock
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @param {string} fromDate Start date in YYYY-MM-DD format
 * @param {string} toDate End date in YYYY-MM-DD format
 * @returns {Promise<string[][]>} Array of sentiment data
 */
export async function getSocialSentiment(ticker, fromDate, toDate) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Validate dates
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fromDate) || !/^\d{4}-\d{2}-\d{2}$/.test(toDate)) {
      return [["Date format must be YYYY-MM-DD"]];
    }

    // Get sentiment data
    const url = `https://api.polygon.io/v2/reference/news?ticker=${ticker}&published_utc.gte=${fromDate}&published_utc.lte=${toDate}&limit=100&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Create header row
      const results = [["Date", "Sentiment Score", "Article Count", "Keywords", "Source"]];
      
      // Group articles by date
      const articlesByDate = new Map();
      
      data.results.forEach(article => {
        const date = article.published_utc.split('T')[0];
        if (!articlesByDate.has(date)) {
          articlesByDate.set(date, []);
        }
        articlesByDate.get(date).push(article);
      });

      // Process each date's articles
      for (const [date, articles] of articlesByDate) {
        // Calculate sentiment score (simple implementation)
        let sentimentScore = 0;
        const keywords = new Set();
        const sources = new Set();
        
        articles.forEach(article => {
          // Add keywords
          if (article.keywords) {
            article.keywords.forEach(keyword => keywords.add(keyword));
          }
          
          // Add sources
          if (article.publisher) {
            sources.add(article.publisher.name);
          }
          
          // Simple sentiment analysis based on title and description
          const text = (article.title + " " + (article.description || "")).toLowerCase();
          const positiveWords = ["up", "rise", "gain", "positive", "bull", "growth", "profit"];
          const negativeWords = ["down", "fall", "loss", "negative", "bear", "decline", "risk"];
          
          positiveWords.forEach(word => {
            if (text.includes(word)) sentimentScore += 1;
          });
          
          negativeWords.forEach(word => {
            if (text.includes(word)) sentimentScore -= 1;
          });
        });

        // Normalize sentiment score
        const normalizedScore = articles.length > 0 ? (sentimentScore / articles.length) : 0;
        
        results.push([
          date,
          normalizedScore.toFixed(2),
          articles.length.toString(),
          Array.from(keywords).slice(0, 3).join(", "), // Show top 3 keywords
          Array.from(sources).slice(0, 2).join(", ") // Show top 2 sources
        ]);
      }

      return results;
    } else {
      return [["No sentiment data found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves insider trading activity
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @param {number} [limit=20] Maximum number of transactions to return
 * @returns {Promise<string[][]>} Array of insider trading data
 */
export async function getInsiderTrading(ticker, limit = 20) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Get insider trading data
    const url = `https://api.polygon.io/v2/reference/insider-transactions?ticker=${ticker}&limit=${limit}&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Create header row
      const results = [["Date", "Insider Name", "Title", "Transaction", "Shares", "Price", "Value"]];
      
      // Process each transaction
      data.results.forEach(transaction => {
        const value = transaction.shares * transaction.price;
        
        results.push([
          transaction.filed_date.split('T')[0],
          transaction.insider_name || "N/A",
          transaction.insider_title || "N/A",
          transaction.transaction_type || "N/A",
          transaction.shares.toLocaleString(),
          "$" + transaction.price.toFixed(2),
          "$" + value.toLocaleString(undefined, {maximumFractionDigits: 0})
        ]);
      });

      // Add summary row
      const totalBuyShares = data.results
        .filter(t => t.transaction_type === "P")
        .reduce((sum, t) => sum + t.shares, 0);
        
      const totalSellShares = data.results
        .filter(t => t.transaction_type === "S")
        .reduce((sum, t) => sum + t.shares, 0);
        
      const netShares = totalBuyShares - totalSellShares;
      
      results.push(
        ["", "", "", "", "", "", ""],
        ["Summary", "", "", "", 
         `Net: ${netShares.toLocaleString()} shares`,
         `Buy: ${totalBuyShares.toLocaleString()}`,
         `Sell: ${totalSellShares.toLocaleString()}`]
      );

      return results;
    } else {
      return [["No insider trading data found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves institutional ownership data
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @returns {Promise<string[][]>} Array of institutional ownership data
 */
export async function getInstitutionalOwnership(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Get institutional holdings data
    const url = `https://api.polygon.io/v2/reference/institutional-holdings?ticker=${ticker}&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Create header rows
      const results = [
        ["Institution", "Shares", "Change", "% of Portfolio", "Market Value"],
        ["==========", "=======", "======", "============", "============"]
      ];
      
      // Sort by number of shares (descending)
      const sortedHoldings = data.results.sort((a, b) => b.shares - a.shares);
      
      // Calculate total shares held by institutions
      const totalShares = sortedHoldings.reduce((sum, holding) => sum + holding.shares, 0);
      
      // Process each institution's holdings
      sortedHoldings.forEach(holding => {
        const marketValue = holding.shares * (holding.average_price || 0);
        const portfolioPercent = (holding.shares / totalShares * 100).toFixed(2) + "%";
        
        results.push([
          holding.institution_name || "N/A",
          holding.shares.toLocaleString(),
          (holding.change_shares > 0 ? "+" : "") + holding.change_shares.toLocaleString(),
          portfolioPercent,
          "$" + marketValue.toLocaleString(undefined, {maximumFractionDigits: 0})
        ]);
      });

      // Add summary rows
      const totalValue = sortedHoldings.reduce((sum, h) => sum + (h.shares * (h.average_price || 0)), 0);
      const averageChange = sortedHoldings.reduce((sum, h) => sum + h.change_shares, 0);
      
      results.push(
        ["", "", "", "", ""],
        ["Total", 
         totalShares.toLocaleString(),
         (averageChange > 0 ? "+" : "") + averageChange.toLocaleString(),
         "100%",
         "$" + totalValue.toLocaleString(undefined, {maximumFractionDigits: 0})]
      );

      return results;
    } else {
      return [["No institutional ownership data found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves economic calendar events
 * @customfunction
 * @param {string} fromDate Start date in YYYY-MM-DD format
 * @param {string} toDate End date in YYYY-MM-DD format
 * @returns {Promise<string[][]>} Array of economic events
 */
export async function getEconomicCalendar(fromDate, toDate) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Validate dates
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fromDate) || !/^\d{4}-\d{2}-\d{2}$/.test(toDate)) {
      return [["Date format must be YYYY-MM-DD"]];
    }

    // Get economic events
    const url = `https://api.polygon.io/v2/reference/calendar?from=${fromDate}&to=${toDate}&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Create header row
      const results = [["Date", "Event", "Actual", "Forecast", "Previous", "Impact"]];
      
      // Sort events by date
      const sortedEvents = data.results.sort((a, b) => 
        new Date(a.date) - new Date(b.date)
      );

      // Process each event
      sortedEvents.forEach(event => {
        results.push([
          event.date,
          event.name || "N/A",
          event.actual || "N/A",
          event.forecast || "N/A",
          event.previous || "N/A",
          event.impact || "N/A"
        ]);
      });

      return results;
    } else {
      return [["No economic events found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Gets the top gainers or losers in the market
 * @customfunction
 * @param {string} [type="gainers"] Type of movers ("gainers" or "losers")
 * @param {number} [limit=20] Maximum number of stocks to return
 * @returns {Promise<string[][]>} Array of top movers
 */
export async function getTopMovers(type = "gainers", limit = 20) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Validate type
    if (!["gainers", "losers"].includes(type.toLowerCase())) {
      return [[`Invalid type. Use "gainers" or "losers"`]];
    }

    // Get snapshot of all tickers
    const url = `https://api.polygon.io/v2/snapshot/locale/us/markets/stocks/tickers?apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.tickers && data.tickers.length > 0) {
      // Calculate percentage changes and filter valid entries
      const tickersWithChanges = data.tickers
        .filter(ticker => 
          ticker.day && 
          ticker.day.o > 0 && 
          ticker.day.c > 0 &&
          ticker.todaysChange !== undefined
        )
        .map(ticker => ({
          symbol: ticker.ticker,
          name: ticker.name || "N/A",
          price: ticker.day.c,
          change: ticker.todaysChange,
          changePercent: ticker.todaysChangePerc,
          volume: ticker.day.v
        }));

      // Sort based on type
      const sortedTickers = type.toLowerCase() === "gainers"
        ? tickersWithChanges.sort((a, b) => b.changePercent - a.changePercent)
        : tickersWithChanges.sort((a, b) => a.changePercent - b.changePercent);

      // Take top N results
      const topMovers = sortedTickers.slice(0, limit);

      // Create results array
      const results = [["Symbol", "Name", "Price", "Change", "Change %", "Volume"]];

      topMovers.forEach(ticker => {
        results.push([
          ticker.symbol,
          ticker.name,
          "$" + ticker.price.toFixed(2),
          (ticker.change >= 0 ? "+" : "") + ticker.change.toFixed(2),
          (ticker.changePercent >= 0 ? "+" : "") + ticker.changePercent.toFixed(2) + "%",
          ticker.volume.toLocaleString()
        ]);
      });

      return results;
    } else {
      return [["No market data found."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves indices data and components
 * @customfunction
 * @param {string[]} indices Array of index symbols (e.g., ["SPX", "NDX", "DJI"])
 * @returns {Promise<string[][]>} Array of index data
 */
export async function getIndicesData(indices) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    if (!Array.isArray(indices) || indices.length === 0) {
      return [["Invalid indices array. Expected format: [SPX, NDX, ...]"]];
    }

    // Create header row
    const results = [["Index", "Price", "Change", "Change %", "Components"]];
    
    // Process each index
    for (const index of indices) {
      try {
        // Get index data
        const url = `https://api.polygon.io/v2/aggs/ticker/I:${index}/prev?adjusted=true&apiKey=${apiKey}`;
        const response = await fetch(url);
        
        if (!response.ok) {
          results.push([index, "Error", "N/A", "N/A", "N/A"]);
          continue;
        }
        
        const data = await response.json();
        if (data.results && data.results.length > 0) {
          const indexData = data.results[0];
          const change = indexData.c - indexData.o;
          const changePercent = (change / indexData.o) * 100;
          
          // Get index components (top 5 by weight)
          const componentsUrl = `https://api.polygon.io/v3/reference/indices/${index}/constituents?apiKey=${apiKey}`;
          const componentsResponse = await fetch(componentsUrl);
          
          let components = "N/A";
          if (componentsResponse.ok) {
            const componentsData = await componentsResponse.json();
            if (componentsData.results) {
              components = componentsData.results
                .slice(0, 5)
                .map(c => c.ticker)
                .join(", ");
            }
          }
          
          results.push([
            index,
            indexData.c.toFixed(2),
            (change >= 0 ? "+" : "") + change.toFixed(2),
            (changePercent >= 0 ? "+" : "") + changePercent.toFixed(2) + "%",
            components
          ]);
        } else {
          results.push([index, "No data", "N/A", "N/A", "N/A"]);
        }
      } catch (error) {
        results.push([index, `Error: ${error.message}`, "N/A", "N/A", "N/A"]);
      }
      
      // Add delay between requests to respect rate limits
      await new Promise(resolve => setTimeout(resolve, 200));
    }
    
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Screens stocks based on technical indicators
 * @customfunction
 * @param {number} minPrice Minimum stock price
 * @param {number} maxPrice Maximum stock price
 * @param {number} minVolume Minimum trading volume
 * @param {number} minChange Minimum price change percentage
 * @param {number} maxChange Maximum price change percentage
 * @param {number} [limit=50] Maximum number of stocks to return
 * @returns {Promise<string[][]>} Array of stocks matching criteria
 */
export async function screenStocks(minPrice, maxPrice, minVolume, minChange, maxChange, limit = 50) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Get snapshot of all tickers
    const url = `https://api.polygon.io/v2/snapshot/locale/us/markets/stocks/tickers?apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data.tickers || data.tickers.length === 0) {
      return [["No market data found."]];
    }

    // Filter tickers based on criteria
    const filteredTickers = data.tickers.filter(ticker => {
      if (!ticker.day || !ticker.prevDay) return false;

      const price = ticker.day.c;
      const volume = ticker.day.v;
      const prevClose = ticker.prevDay.c;
      const changePercent = ((price - prevClose) / prevClose) * 100;

      // Apply screening criteria
      if (minPrice && price < minPrice) return false;
      if (maxPrice && price > maxPrice) return false;
      if (minVolume && volume < minVolume) return false;
      if (minChange && changePercent < minChange) return false;
      if (maxChange && changePercent > maxChange) return false;

      return true;
    });

    // Sort results by volume (descending)
    const sortedTickers = filteredTickers
      .sort((a, b) => b.day.v - a.day.v)
      .slice(0, limit);

    // Create results array
    const results = [["Symbol", "Price", "Change %", "Volume", "Market Cap"]];

    sortedTickers.forEach(ticker => {
      const price = ticker.day.c;
      const prevClose = ticker.prevDay.c;
      const changePercent = ((price - prevClose) / prevClose) * 100;
      const marketCap = ticker.market_cap || "N/A";

      results.push([
        ticker.ticker,
        "$" + price.toFixed(2),
        (changePercent >= 0 ? "+" : "") + changePercent.toFixed(2) + "%",
        ticker.day.v.toLocaleString(),
        typeof marketCap === "number" 
          ? "$" + (marketCap / 1e6).toFixed(0) + "M"
          : marketCap
      ]);
    });

    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Calculates pivot points using various methods
 * @customfunction
 * @param {string} ticker The stock ticker symbol
 * @param {string} [method="standard"] Pivot point method ("standard", "woodie", "camarilla", "fibonacci")
 * @returns {Promise<string[][]>} Array of pivot points and support/resistance levels
 */
export async function getPivotPoints(ticker, method = "standard") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Validate method
    const validMethods = ["standard", "woodie", "camarilla", "fibonacci"];
    if (!validMethods.includes(method.toLowerCase())) {
      return [[`Invalid method. Valid options are: ${validMethods.join(", ")}`]];
    }

    // Get previous day's data
    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/prev?adjusted=true&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data.results || data.results.length === 0) {
      return [["No price data found."]];
    }

    const bar = data.results[0];
    const high = bar.h;
    const low = bar.l;
    const close = bar.c;
    const open = bar.o;

    let pivotPoints;
    switch (method.toLowerCase()) {
      case "standard": {
        const pivot = (high + low + close) / 3;
        pivotPoints = {
          r3: pivot + ((high - low) * 2),
          r2: pivot + (high - low),
          r1: (2 * pivot) - low,
          pivot,
          s1: (2 * pivot) - high,
          s2: pivot - (high - low),
          s3: pivot - ((high - low) * 2)
        };
        break;
      }

      case "woodie": {
        const pivot = (high + low + (2 * close)) / 4;
        pivotPoints = {
          r3: high + 2 * (pivot - low),
          r2: pivot + (high - low),
          r1: (2 * pivot) - low,
          pivot,
          s1: (2 * pivot) - high,
          s2: pivot - (high - low),
          s3: low - 2 * (high - pivot)
        };
        break;
      }

      case "camarilla": {
        pivotPoints = {
          r4: close + ((high - low) * 1.5),
          r3: close + ((high - low) * 1.25),
          r2: close + ((high - low) * 1.1666),
          r1: close + ((high - low) * 1.0833),
          pivot: (high + low + close) / 3,
          s1: close - ((high - low) * 1.0833),
          s2: close - ((high - low) * 1.1666),
          s3: close - ((high - low) * 1.25),
          s4: close - ((high - low) * 1.5)
        };
        break;
      }

      case "fibonacci": {
        const pivot = (high + low + close) / 3;
        const range = high - low;
        pivotPoints = {
          r3: pivot + (range * 1.000),
          r2: pivot + (range * 0.618),
          r1: pivot + (range * 0.382),
          pivot,
          s1: pivot - (range * 0.382),
          s2: pivot - (range * 0.618),
          s3: pivot - (range * 1.000)
        };
        break;
      }
    }

    // Create results array
    const results = [
      ["Level", "Price", "Distance %"],
      ["=====", "=====", "========="]
    ];

    // Add pivot points to results
    Object.entries(pivotPoints).forEach(([level, price]) => {
      const distance = ((price - close) / close) * 100;
      results.push([
        level.toUpperCase(),
        "$" + price.toFixed(2),
        (distance >= 0 ? "+" : "") + distance.toFixed(2) + "%"
      ]);
    });

    // Add current price for reference
    results.push(
      ["", "", ""],
      ["Current", "$" + close.toFixed(2), ""]
    );

    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Gets previous close data for multiple tickers
 * @customfunction
 * @param {string[]} tickers Array of stock ticker symbols
 * @returns {Promise<string[][]>} Array of previous close data
 */
export async function getPreviousClose(tickers) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    if (!Array.isArray(tickers) || tickers.length === 0) {
      return [["Invalid tickers array. Expected format: [AAPL, MSFT, ...]"]];
    }

    // Create header row
    const results = [["Symbol", "Close", "Open", "High", "Low", "Volume", "VWAP"]];
    
    // Process tickers in batches of 5 to avoid rate limits
    for (let i = 0; i < tickers.length; i += 5) {
      const batch = tickers.slice(i, i + 5);
      const promises = batch.map(ticker => 
        fetch(`https://api.polygon.io/v2/aggs/ticker/${ticker}/prev?adjusted=true&apiKey=${apiKey}`)
          .then(response => response.json())
      );

      const responses = await Promise.all(promises);
      
      responses.forEach((data, index) => {
        if (data.results && data.results.length > 0) {
          const bar = data.results[0];
          
          results.push([
            batch[index],
            "$" + bar.c.toFixed(2),
            "$" + bar.o.toFixed(2),
            "$" + bar.h.toFixed(2),
            "$" + bar.l.toFixed(2),
            bar.v.toLocaleString(),
            "$" + (bar.vw || 0).toFixed(2)
          ]);
        } else {
          results.push([batch[index], "No data", "N/A", "N/A", "N/A", "N/A", "N/A"]);
        }
      });

      // Add delay between batches to respect rate limits
      if (i + 5 < tickers.length) {
        await new Promise(resolve => setTimeout(resolve, 200));
      }
    }
    
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Gets grouped daily bars for the entire market
 * @customfunction
 * @param {string} date Date in YYYY-MM-DD format
 * @param {string} [locale="us"] Market locale
 * @param {string} [market="stocks"] Market type ("stocks", "crypto", "fx")
 * @returns {Promise<string[][]>} Array of daily bars data
 */
export async function getGroupedDailyBars(date, locale = "us", market = "stocks") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Validate date
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      return [["Date format must be YYYY-MM-DD"]];
    }

    // Validate market
    const validMarkets = ["stocks", "crypto", "fx"];
    if (!validMarkets.includes(market.toLowerCase())) {
      return [[`Invalid market. Valid options are: ${validMarkets.join(", ")}`]];
    }

    const url = `https://api.polygon.io/v2/aggs/grouped/locale/${locale}/market/${market}/${date}?adjusted=true&apiKey=${apiKey}`;
    const response = await fetch(url);

    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (data.results && data.results.length > 0) {
      // Create header row
      const results = [["Symbol", "Open", "High", "Low", "Close", "Volume", "VWAP", "Change %"]];
      
      // Sort by volume (descending)
      const sortedBars = data.results
        .filter(bar => bar.v > 0) // Filter out zero volume
        .sort((a, b) => b.v - a.v);

      // Process each bar
      sortedBars.forEach(bar => {
        const changePercent = ((bar.c / bar.o) - 1) * 100;
        
        results.push([
          bar.T || "N/A", // Symbol
          "$" + bar.o.toFixed(2),
          "$" + bar.h.toFixed(2),
          "$" + bar.l.toFixed(2),
          "$" + bar.c.toFixed(2),
          bar.v.toLocaleString(),
          "$" + (bar.vw || 0).toFixed(2),
          (changePercent >= 0 ? "+" : "") + changePercent.toFixed(2) + "%"
        ]);
      });

      // Add summary row
      const totalVolume = sortedBars.reduce((sum, bar) => sum + bar.v, 0);
      const avgChange = sortedBars.reduce((sum, bar) => sum + ((bar.c / bar.o) - 1), 0) / sortedBars.length * 100;
      
      results.push(
        ["", "", "", "", "", "", "", ""],
        ["Summary",
         `${sortedBars.length} symbols`,
         "",
         "",
         "",
         `Vol: ${totalVolume.toLocaleString()}`,
         "",
         `Avg: ${avgChange.toFixed(2)}%`]
      );

      return results;
    } else {
      return [["No data found for the specified date."]];
    }
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}
