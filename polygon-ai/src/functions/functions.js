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

    limit = Math.min(Math.max(1, limit), 50);
    const url = `https://api.polygon.io/v3/reference/tickers?search=${encodeURIComponent(searchTerm)}&active=true&limit=${limit}&apiKey=${apiKey}`;
    
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data.results) {
      return [["No data returned from API."]];
    }
    if (data.results.length > 0) {
      const results = [["Symbol", "Name", "Market", "Type", "Primary Exchange"]];
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
    if (!data.results) {
      return [["No data returned from API."]];
    }
    if (data.results.length > 0) {
      const results = [["Ticker", "Ex-Date", "Payment Date", "Ratio", "To Factor", "From Factor"]];
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

    if (!/^\d{4}-\d{2}-\d{2}$/.test(fromDate) || !/^\d{4}-\d{2}-\d{2}$/.test(toDate)) {
      return [["Date format must be YYYY-MM-DD"]];
    }

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
    if (!data.results) {
      return [["No data returned from API."]];
    }
    if (data.results.length > 0) {
      const results = [["Date", "Open", "High", "Low", "Close", "Volume", "VWAP"]];
      data.results.forEach(bar => {
        const date = new Date(bar.t).toISOString().split('T')[0];
        results.push([
          date,
          bar.o?.toFixed(2) || "N/A",
          bar.h?.toFixed(2) || "N/A",
          bar.l?.toFixed(2) || "N/A",
          bar.c?.toFixed(2) || "N/A",
          bar.v || "N/A",
          bar.vw?.toFixed(2) || "N/A"
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
    if (!data.results) {
      return [["No data returned from API."]];
    }
    if (data.results.length > 0) {
      const filteredExchanges = data.results.filter(exchange => 
        exchangeType === "exchange" ? !exchange.type.includes("otc") && !exchange.type.includes("index") :
        exchange.type.includes(exchangeType)
      );

      if (filteredExchanges.length === 0) {
        return [[`No ${exchangeType} exchanges found.`]];
      }

      const results = [["Name", "Market Identifier Code (MIC)", "Type", "Market", "Country", "Operating MIC"]];
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

    if (!year) {
      year = new Date().getFullYear();
    }
    
    const url = `https://api.polygon.io/v1/marketstatus/upcoming?apiKey=${apiKey}`;
    
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const holidays = await response.json();
    if (!holidays) {
      return [["No data returned from API."]];
    }
    if (holidays.length > 0) {
      const filteredHolidays = holidays.filter(holiday => {
        const holidayDate = new Date(holiday.date);
        return holidayDate.getFullYear() === year;
      });

      if (filteredHolidays.length === 0) {
        return [[`No market holidays found for ${year}.`]];
      }

      const results = [["Date", "Holiday", "Status", "Open", "Close"]];
      filteredHolidays.forEach(holiday => {
        const date = new Date(holiday.date).toISOString().split('T')[0];
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
 * @note Approximates sector performance using sector ETFs
 */
export async function getSectorPerformance(timespan = "day") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

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

    const spyTicker = "SPY";
    const endDate = new Date();
    const startDate = new Date();
    
    switch(timespan) {
      case "day": startDate.setDate(startDate.getDate() - 1); break;
      case "week": startDate.setDate(startDate.getDate() - 7); break;
      case "month": startDate.setMonth(startDate.getMonth() - 1); break;
      case "quarter": startDate.setMonth(startDate.getMonth() - 3); break;
      case "year": startDate.setFullYear(startDate.getFullYear() - 1); break;
      default: startDate.setDate(startDate.getDate() - 1);
    }

    const fromDate = startDate.toISOString().split('T')[0];
    const toDate = endDate.toISOString().split('T')[0];

    const spyUrl = `https://api.polygon.io/v2/aggs/ticker/${spyTicker}/range/1/day/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const spyResponse = await fetch(spyUrl);
    if (!spyResponse.ok) {
      return [[`HTTP error! status: ${spyResponse.status}`]];
    }
    
    const spyData = await spyResponse.json();
    if (!spyData.results || spyData.results.length < 2) {
      return [["Insufficient S&P 500 data for the selected timespan."]];
    }

    const spyStartPrice = spyData.results[0].c;
    const spyEndPrice = spyData.results[spyData.results.length - 1].c;
    const spyPerformance = ((spyEndPrice / spyStartPrice) - 1) * 100;

    const results = [["Sector", "Performance (%)", "Relative to S&P 500", "Ticker"]];
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
    
    results.sort((a, b) => {
      if (a[0] === "Sector") return -1;
      if (b[0] === "Sector") return 1;
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

    days = Math.min(Math.max(5, days), 365);
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - (days * 2));
    
    const fromDate = startDate.toISOString().split('T')[0];
    const toDate = endDate.toISOString().split('T')[0];

    const responses = await Promise.all([
      fetch(`https://api.polygon.io/v2/aggs/ticker/${ticker1}/range/1/day/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`),
      fetch(`https://api.polygon.io/v2/aggs/ticker/${ticker2}/range/1/day/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`)
    ]);

    if (!responses[0].ok || !responses[1].ok) {
      return `HTTP error! status: ${responses[0].status || responses[1].status}`;
    }

    const data = await Promise.all([responses[0].json(), responses[1].json()]);
    if (!data[0].results || !data[1].results || 
        data[0].results.length < days || data[1].results.length < days) {
      return "Insufficient price data for correlation calculation.";
    }

    const prices1 = data[0].results.slice(-days).map(bar => bar.c);
    const prices2 = data[1].results.slice(-days).map(bar => bar.c);

    const returns1 = [];
    const returns2 = [];
    for (let i = 1; i < prices1.length; i++) {
      returns1.push((prices1[i] / prices1[i-1]) - 1);
      returns2.push((prices2[i] / prices2[i-1]) - 1);
    }

    const correlation = calculateCorrelation(returns1, returns2);
    return parseFloat(correlation.toFixed(4));
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
  const xMean = x.reduce((a, b) => a + b, 0) / n;
  const yMean = y.reduce((a, b) => a + b, 0) / n;
  
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

    if (!Array.isArray(portfolioData) || !portfolioData.every(row => Array.isArray(row) && row.length >= 2)) {
      return [["Invalid portfolio data. Expected format: [[ticker, shares], ...]"]];
    }

    const results = [["Symbol", "Shares", "Current Price", "Market Value", "Day Change %"]];
    let totalValue = 0;
    
    for (const position of portfolioData) {
      const [ticker, sharesInput] = position;
      const shares = parseFloat(sharesInput);
      
      if (isNaN(shares) || shares <= 0) {
        results.push([ticker, "Invalid shares", "N/A", "N/A", "N/A"]);
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
        const previousClose = priceData.o;
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
    const apiKey = getApiKey();
    if (!apiKey) {
      return "API key not set. Please set your Polygon.io API key.";
    }

    const url = `https://api.polygon.io/v3/reference/tickers?ticker=${ticker}&active=true&limit=100&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) {
      return `HTTP error! status: ${response.status}`;
    }

    const data = await response.json();
    if (!data.results) {
      return "No data returned from API.";
    }
    if (data.results.length > 0) {
      const details = data.results[0];
      if (property) {
        return details[property] || "Property not found.";
      } else {
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
      return [["API key not set. Please set your Polygon.io API key."]];
    }
    limit = Math.max(1, Math.min(limit, 50));
    const url = `https://api.polygon.io/v2/reference/news?ticker=${ticker}&limit=${limit}&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data.results) {
      return [["No data returned from API."]];
    }
    if (data.results.length > 0) {
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
    if (!data.results) {
      return "No data returned from API.";
    }
    if (data.results.length > 0) {
      const priceData = data.results[0];
      if (property && priceData[property] !== undefined) {
        return priceData[property];
      } else if (property) {
        return "Property not found.";
      } else {
        return priceData.c;
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
    if (!data) {
      return "No data returned from API.";
    }
    
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
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    limit = Math.max(1, Math.min(limit, 50));
    const url = `https://api.polygon.io/v3/reference/dividends?ticker=${ticker}&limit=${limit}&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data.results) {
      return [["No data returned from API."]];
    }
    if (data.results.length > 0) {
      const results = [["Ex-Dividend Date", "Payment Date", "Record Date", "Cash Amount", "Declaration Date", "Frequency"]];
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
 * @param {number} [period=14] The period to use for SMA, EMA, or RSI calculation.
 * @param {string} [from] The start date in YYYY-MM-DD format. Defaults to 30 days ago.
 * @param {string} [to] The end date in YYYY-MM-DD format. Defaults to today.
 * @returns {string[][]} A 2D array of indicator values with dates or an error message.
 * @note MACD uses fixed periods (12, 26, 9) and ignores the period parameter.
 */
export async function getTechnicalIndicator(ticker, indicator, period = 14, from, to) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    if (!ticker) {
      return [["Missing required parameter: ticker"]];
    }
    if (!indicator) {
      return [["Missing required parameter: indicator"]];
    }

    if (!from) {
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
      from = thirtyDaysAgo.toISOString().split('T')[0];
    }
    if (!to) {
      to = new Date().toISOString().split('T')[0];
    }

    if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) {
      return [["Date format must be YYYY-MM-DD"]];
    }

    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/day/${from}/${to}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data.results || data.results.length === 0) {
      return [["Insufficient price data to calculate indicator."]];
    }
    
    const prices = data.results.map(bar => bar.c);
    const dates = data.results.map(bar => new Date(bar.t).toISOString().split('T')[0]);
    
    if (!prices.every(price => typeof price === 'number' && !isNaN(price))) {
      return [["Invalid price data received"]];
    }

    if (period <= 0 || period >= prices.length) {
      return [[`Invalid period: ${period}. Must be between 1 and ${prices.length-1}`]];
    }

    let result = [];
    let indicatorValues = [];
    let startIndex = 0;
    
    switch (indicator.toUpperCase()) {
      case "SMA":
        indicatorValues = calculateSMA(prices, period);
        startIndex = period - 1;
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
        const macdResult = calculateMACD(prices, 12, 26, 9);
        const [macdLine, signalLine, histogram] = macdResult;
        startIndex = prices.length - macdLine.length;
        result.push(["Date", "MACD Line", "Signal Line", "Histogram"]);
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
    
    for (let i = 0; i < indicatorValues.length; i++) {
      const dateIndex = startIndex + i;
      if (dateIndex < dates.length) {
        result.push([dates[dateIndex], indicatorValues[i].toFixed(4)]);
      }
    }
    
    return result;
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
  if (!Array.isArray(prices) || prices.length === 0) throw new Error("Invalid prices array");
  if (period <= 0 || period > prices.length) throw new Error(`Invalid period: ${period}`);
  
  const result = [];
  for (let i = period - 1; i < prices.length; i++) {
    let sum = 0;
    for (let j = 0; j < period; j++) {
      const price = prices[i - j];
      if (typeof price !== 'number' || isNaN(price)) throw new Error(`Invalid price at position ${i-j}: ${price}`);
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
  if (!Array.isArray(prices) || prices.length === 0) throw new Error("Invalid prices array");
  if (period <= 0 || period > prices.length) throw new Error(`Invalid period: ${period}`);
  
  const result = [];
  const multiplier = 2 / (period + 1);
  
  for (let i = 0; i < prices.length; i++) {
    if (typeof prices[i] !== 'number' || isNaN(prices[i])) throw new Error(`Invalid price at position ${i}: ${prices[i]}`);
  }
  
  let ema = prices.slice(0, period).reduce((a, b) => a + b, 0) / period;
  result.push(ema);
  
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
  if (!Array.isArray(prices) || prices.length === 0) throw new Error("Invalid prices array");
  if (period <= 0 || period > prices.length - 1) throw new Error(`Invalid period: ${period}`);
  
  const result = [];
  const gains = [];
  const losses = [];
  
  for (let i = 1; i < prices.length; i++) {
    if (typeof prices[i] !== 'number' || isNaN(prices[i]) || typeof prices[i-1] !== 'number' || isNaN(prices[i-1])) {
      throw new Error(`Invalid price at positions ${i-1} or ${i}`);
    }
    const change = prices[i] - prices[i-1];
    gains.push(change > 0 ? change : 0);
    losses.push(change < 0 ? -change : 0);
  }
  
  for (let i = period; i < gains.length; i++) {
    let avgGain = gains.slice(i - period, i).reduce((a, b) => a + b, 0) / period;
    let avgLoss = losses.slice(i - period, i).reduce((a, b) => a + b, 0) / period;
    if (avgLoss === 0) {
      result.push(100);
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
  if (!Array.isArray(prices) || prices.length === 0) throw new Error("Invalid prices array");
  if (fastPeriod <= 0 || slowPeriod <= 0 || signalPeriod <= 0) throw new Error("Periods must be positive numbers");
  if (slowPeriod <= fastPeriod) throw new Error("Slow period must be greater than fast period");
  if (prices.length <= slowPeriod) throw new Error(`Not enough price data for the specified periods. Need at least ${slowPeriod + signalPeriod} points.`);
  
  const fastEMA = calculateEMA(prices, fastPeriod);
  const slowEMA = calculateEMA(prices, slowPeriod);
  
  const macdLine = [];
  const offset = slowPeriod - fastPeriod;
  for (let i = 0; i < slowEMA.length; i++) {
    macdLine.push(fastEMA[i + offset] - slowEMA[i]);
  }
  
  if (macdLine.length < signalPeriod) throw new Error("Insufficient data to calculate signal line");
  
  const signalLine = calculateEMA(macdLine, signalPeriod);
  const histogram = [];
  const histOffset = macdLine.length - signalLine.length;
  
  for (let i = 0; i < signalLine.length; i++) {
    histogram.push(macdLine[i + histOffset] - signalLine[i]);
  }
  
  const minLength = Math.min(histogram.length, signalLine.length);
  const macdLineAligned = macdLine.slice(macdLine.length - minLength);
  
  return [macdLineAligned, signalLine.slice(0, minLength), histogram.slice(0, minLength)];
}

/**
 * Calculates Bollinger Bands for a stock
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL")
 * @param {number} [period=20] Period for SMA calculation
 * @param {number} [stdDev=2] Number of standard deviations
 * @param {string} [from] Start date in YYYY-MM-DD format
 * @param {string} [to] End date in YYYY-MM-DD format
 * @returns {Promise<string[][]>} Array containing dates and Bollinger Bands values
 */
export async function getBollingerBands(ticker, period = 20, stdDev = 2, from, to) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    if (!from) {
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
      from = thirtyDaysAgo.toISOString().split('T')[0];
    }
    if (!to) {
      to = new Date().toISOString().split('T')[0];
    }

    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/day/${from}/${to}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data.results || data.results.length === 0) {
      return [["Insufficient price data to calculate Bollinger Bands."]];
    }

    const prices = data.results.map(bar => bar.c);
    const dates = data.results.map(bar => new Date(bar.t).toISOString().split('T')[0]);
    
    const results = [["Date", "Middle Band (SMA)", "Upper Band", "Lower Band"]];
    
    for (let i = period - 1; i < prices.length; i++) {
      const slice = prices.slice(i - period + 1, i + 1);
      const sma = slice.reduce((a, b) => a + b) / period;
      
      const squaredDiffs = slice.map(price => Math.pow(price - sma, 2));
      const variance = squaredDiffs.reduce((a, b) => a + b) / period;
      const standardDeviation = Math.sqrt(variance);
      
      const upperBand = sma + (standardDeviation * stdDev);
      const lowerBand = sma - (standardDeviation * stdDev);
      
      results.push([
        dates[i],
        sma.toFixed(4),
        upperBand.toFixed(4),
        lowerBand.toFixed(4)
      ]);
    }
    
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Calculates Average True Range (ATR) for a stock
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL")
 * @param {number} [period=14] Period for ATR calculation
 * @param {string} [from] Start date in YYYY-MM-DD format
 * @param {string} [to] End date in YYYY-MM-DD format
 * @returns {Promise<string[][]>} Array containing dates and ATR values
 */
export async function getATR(ticker, period = 14, from, to) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    if (!from) {
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
      from = thirtyDaysAgo.toISOString().split('T')[0];
    }
    if (!to) {
      to = new Date().toISOString().split('T')[0];
    }

    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/day/${from}/${to}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data.results || data.results.length === 0) {
      return [["Insufficient price data to calculate ATR."]];
    }

    const results = [["Date", "ATR"]];
    const bars = data.results;
    
    // Calculate True Range series
    const trueRanges = [];
    for (let i = 1; i < bars.length; i++) {
      const high = bars[i].h;
      const low = bars[i].l;
      const prevClose = bars[i-1].c;
      
      const tr1 = high - low;
      const tr2 = Math.abs(high - prevClose);
      const tr3 = Math.abs(low - prevClose);
      
      const trueRange = Math.max(tr1, tr2, tr3);
      trueRanges.push(trueRange);
    }
    
    // Calculate ATR
    let atr = trueRanges.slice(0, period).reduce((a, b) => a + b) / period;
    results.push([bars[period].t, atr.toFixed(4)]);
    
    for (let i = period; i < trueRanges.length; i++) {
      atr = ((atr * (period - 1)) + trueRanges[i]) / period;
      const date = new Date(bars[i + 1].t).toISOString().split('T')[0];
      results.push([date, atr.toFixed(4)]);
    }
    
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Calculates Pivot Points for a stock
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL")
 * @param {string} [method="standard"] Pivot point calculation method ("standard", "fibonacci", "woodie", "demark")
 * @returns {Promise<string[][]>} Array containing pivot points and support/resistance levels
 */
export async function getPivotPoints(ticker, method = "standard") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    // Get previous day's data
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const from = yesterday.toISOString().split('T')[0];
    const to = new Date().toISOString().split('T')[0];

    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/day/${from}/${to}?adjusted=true&sort=desc&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data.results || data.results.length === 0) {
      return [["Insufficient price data to calculate pivot points."]];
    }

    const bar = data.results[0];
    const high = bar.h;
    const low = bar.l;
    const close = bar.c;
    
    let results = [["Level", "Value"]];
    
    switch (method.toLowerCase()) {
      case "standard": {
        const pp = (high + low + close) / 3;
        const r1 = (2 * pp) - low;
        const s1 = (2 * pp) - high;
        const r2 = pp + (high - low);
        const s2 = pp - (high - low);
        const r3 = high + 2 * (pp - low);
        const s3 = low - 2 * (high - pp);
        
        results.push(
          ["Pivot Point", pp.toFixed(4)],
          ["Resistance 1", r1.toFixed(4)],
          ["Resistance 2", r2.toFixed(4)],
          ["Resistance 3", r3.toFixed(4)],
          ["Support 1", s1.toFixed(4)],
          ["Support 2", s2.toFixed(4)],
          ["Support 3", s3.toFixed(4)]
        );
        break;
      }
      case "fibonacci": {
        const pp = (high + low + close) / 3;
        const r1 = pp + 0.382 * (high - low);
        const r2 = pp + 0.618 * (high - low);
        const r3 = pp + 1.000 * (high - low);
        const s1 = pp - 0.382 * (high - low);
        const s2 = pp - 0.618 * (high - low);
        const s3 = pp - 1.000 * (high - low);
        
        results.push(
          ["Pivot Point", pp.toFixed(4)],
          ["Resistance 1 (38.2%)", r1.toFixed(4)],
          ["Resistance 2 (61.8%)", r2.toFixed(4)],
          ["Resistance 3 (100%)", r3.toFixed(4)],
          ["Support 1 (38.2%)", s1.toFixed(4)],
          ["Support 2 (61.8%)", s2.toFixed(4)],
          ["Support 3 (100%)", s3.toFixed(4)]
        );
        break;
      }
      case "woodie": {
        const pp = (high + low + 2 * close) / 4;
        const r1 = 2 * pp - low;
        const r2 = pp + high - low;
        const s1 = 2 * pp - high;
        const s2 = pp - high + low;
        
        results.push(
          ["Pivot Point", pp.toFixed(4)],
          ["Resistance 1", r1.toFixed(4)],
          ["Resistance 2", r2.toFixed(4)],
          ["Support 1", s1.toFixed(4)],
          ["Support 2", s2.toFixed(4)]
        );
        break;
      }
      case "demark": {
        let x;
        if (close < bar.o) x = high + (2 * low) + close;
        else if (close > bar.o) x = (2 * high) + low + close;
        else x = high + low + (2 * close);
        
        const pp = x / 4;
        const r1 = x / 2 - low;
        const s1 = x / 2 - high;
        
        results.push(
          ["Pivot Point", pp.toFixed(4)],
          ["Resistance 1", r1.toFixed(4)],
          ["Support 1", s1.toFixed(4)]
        );
        break;
      }
      default:
        return [["Invalid pivot point calculation method. Valid options: standard, fibonacci, woodie, demark"]];
    }
    
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves the daily open and close prices for a ticker on a specific date.
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @param {string} date The date in YYYY-MM-DD format.
 * @returns {Promise<string[][]>} A 2D array with the open, close, afterHours, preMarket, and status.
 */
export async function getDailyOpenClose(ticker, date) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      return [["Invalid date format. Use YYYY-MM-DD."]];
    }

    const url = `https://api.polygon.io/v1/open-close/${ticker}/${date}?apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data) {
      return [["No data returned from API."]];
    }

    return [
      ["Status", "Open", "Close", "After Hours", "Pre Market"],
      [
        data.status || "N/A",
        data.open?.toFixed(2) || "N/A",
        data.close?.toFixed(2) || "N/A",
        data.afterHours?.toFixed(2) || "N/A",
        data.preMarket?.toFixed(2) || "N/A"
      ]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves the last trade for a ticker.
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @returns {Promise<string[][]>} A 2D array with the last trade details.
 */
export async function getLastTrade(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    const url = `https://api.polygon.io/v2/last/trade/${ticker}?apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data.results) {
      return [["No data returned from API."]];
    }

    const trade = data.results;
    return [
      ["Price", "Size", "Exchange", "Conditions", "Timestamp"],
      [
        trade.p?.toFixed(2) || "N/A",
        trade.s || "N/A",
        trade.x || "N/A",
        trade.c?.join(", ") || "N/A",
        new Date(trade.t).toLocaleString() || "N/A"
      ]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves the last quote for a ticker.
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @returns {Promise<string[][]>} A 2D array with the last quote details.
 */
export async function getLastQuote(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    const url = `https://api.polygon.io/v2/last/nbbo/${ticker}?apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data.results) {
      return [["No data returned from API."]];
    }

    const quote = data.results;
    return [
      ["Bid Price", "Bid Size", "Ask Price", "Ask Size", "Timestamp"],
      [
        quote.P?.toFixed(2) || "N/A",
        quote.S || "N/A",
        quote.p?.toFixed(2) || "N/A",
        quote.s || "N/A",
        new Date(quote.t).toLocaleString() || "N/A"
      ]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves a snapshot of a specific ticker.
 * @customfunction
 * @param {string} ticker The stock ticker symbol (e.g., "AAPL").
 * @returns {Promise<string[][]>} A 2D array with snapshot data.
 */
export async function getSnapshotTicker(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return [["API key not set. Please set your Polygon.io API key."]];
    }

    const url = `https://api.polygon.io/v2/snapshot/locale/us/markets/stocks/tickers/${ticker}?apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) {
      return [[`HTTP error! status: ${response.status}`]];
    }

    const data = await response.json();
    if (!data.ticker) {
      return [["No data returned from API."]];
    }

    const snapshot = data.ticker;
    const lastTrade = snapshot.lastTrade || {};
    const lastQuote = snapshot.lastQuote || {};
    const dailyBar = snapshot.day || {};

    return [
      ["Last Trade Price", "Last Bid", "Last Ask", "Daily Change", "Daily Change %"],
      [
        lastTrade.p?.toFixed(2) || "N/A",
        lastQuote.P?.toFixed(2) || "N/A",
        lastQuote.p?.toFixed(2) || "N/A",
        (dailyBar.c - dailyBar.o)?.toFixed(2) || "N/A",
        (((dailyBar.c - dailyBar.o) / dailyBar.o) * 100)?.toFixed(2) + "%" || "N/A"
      ]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

// This file contains custom functions for accessing various Polygon.io APIs, including Options, Indices, Forex, Crypto, and Reference data.
// Each function is designed to be used in an Excel add-in and returns data in a 2D array format suitable for spreadsheets.


function isValidDate(date) {
  return /^\d{4}-\d{2}-\d{2}$/.test(date);
}

// ---------------------------
// Options APIs
// ---------------------------

/**
 * Retrieves historical OHLC data for an options contract.
 * @customfunction
 * @param {string} ticker The options contract ticker (e.g., "O:AAPL241025C00250000").
 * @param {string} fromDate Start date in YYYY-MM-DD format.
 * @param {string} toDate End date in YYYY-MM-DD format.
 * @param {string} [timespan="day"] Timespan ("minute", "hour", "day").
 * @returns {Promise<string[][]>} Array of OHLC data.
 */
export async function getOptionsHistoricalOHLC(ticker, fromDate, toDate, timespan = "day") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set. Please set your Polygon.io API key."]];
    if (!isValidDate(fromDate) || !isValidDate(toDate)) return [["Invalid date format. Use YYYY-MM-DD."]];
    
    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/${timespan}/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned from API."]];
    
    const results = [["Date", "Open", "High", "Low", "Close", "Volume"]];
    data.results.forEach(bar => {
      const date = new Date(bar.t).toISOString().split('T')[0];
      results.push([date, bar.o, bar.h, bar.l, bar.c, bar.v]);
    });
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves the last trade for an options contract.
 * @customfunction
 * @param {string} ticker The options contract ticker.
 * @returns {Promise<string[][]>} Array with last trade details.
 */
export async function getOptionsLastTrade(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v2/last/trade/${ticker}?apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const trade = data.results;
    return [
      ["Price", "Size", "Timestamp"],
      [trade.p, trade.s, new Date(trade.t).toISOString()]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves the last quote for an options contract.
 * @customfunction
 * @param {string} ticker The options contract ticker.
 * @returns {Promise<string[][]>} Array with last quote details.
 */
export async function getOptionsLastQuote(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v2/last/nbbo/${ticker}?apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const quote = data.results;
    return [
      ["Bid", "Ask", "Timestamp"],
      [quote.bp, quote.ap, new Date(quote.t).toISOString()]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves a snapshot of an options contract.
 * @customfunction
 * @param {string} ticker The options contract ticker.
 * @returns {Promise<string[][]>} Array with snapshot data.
 */
export async function getOptionsSnapshot(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v3/snapshot?ticker.any=${ticker}&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results || data.results.length === 0) return [["No data returned."]];
    
    const snap = data.results[0];
    return [
      ["Last Price", "Open", "High", "Low", "Volume"],
      [snap.lastTrade.p, snap.day.o, snap.day.h, snap.day.l, snap.day.v]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

// ---------------------------
// Indices APIs
// ---------------------------

/**
 * Retrieves historical OHLC data for an index.
 * @customfunction
 * @param {string} ticker The index ticker (e.g., "I:SPX").
 * @param {string} fromDate Start date in YYYY-MM-DD format.
 * @param {string} toDate End date in YYYY-MM-DD format.
 * @param {string} [timespan="day"] Timespan ("minute", "hour", "day").
 * @returns {Promise<string[][]>} Array of OHLC data.
 */
export async function getIndexHistoricalOHLC(ticker, fromDate, toDate, timespan = "day") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    if (!isValidDate(fromDate) || !isValidDate(toDate)) return [["Invalid date format."]];
    
    const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/${timespan}/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const results = [["Date", "Open", "High", "Low", "Close", "Volume"]];
    data.results.forEach(bar => {
      const date = new Date(bar.t).toISOString().split('T')[0];
      results.push([date, bar.o, bar.h, bar.l, bar.c, bar.v]);
    });
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves a snapshot of an index.
 * @customfunction
 * @param {string} ticker The index ticker.
 * @returns {Promise<string[][]>} Array with snapshot data.
 */
export async function getIndexSnapshot(ticker) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v3/snapshot?ticker.any=${ticker}&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results || data.results.length === 0) return [["No data returned."]];
    
    const snap = data.results[0];
    return [
      ["Last Price", "Open", "High", "Low", "Volume"],
      [snap.lastTrade.p, snap.day.o, snap.day.h, snap.day.l, snap.day.v]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

// ---------------------------
// Forex APIs
// ---------------------------

/**
 * Retrieves historical OHLC data for a forex pair.
 * @customfunction
 * @param {string} pair The forex pair (e.g., "C:EURUSD").
 * @param {string} fromDate Start date in YYYY-MM-DD format.
 * @param {string} toDate End date in YYYY-MM-DD format.
 * @param {string} [timespan="day"] Timespan ("minute", "hour", "day").
 * @returns {Promise<string[][]>} Array of OHLC data.
 */
export async function getForexHistoricalOHLC(pair, fromDate, toDate, timespan = "day") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    if (!isValidDate(fromDate) || !isValidDate(toDate)) return [["Invalid date format."]];
    
    const url = `https://api.polygon.io/v2/aggs/ticker/${pair}/range/1/${timespan}/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const results = [["Date", "Open", "High", "Low", "Close"]];
    data.results.forEach(bar => {
      const date = new Date(bar.t).toISOString().split('T')[0];
      results.push([date, bar.o, bar.h, bar.l, bar.c]);
    });
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves grouped daily OHLC for all forex pairs.
 * @customfunction
 * @param {string} date Date in YYYY-MM-DD format.
 * @returns {Promise<string[][]>} Array of daily OHLC data.
 */
export async function getForexGroupedDailyBars(date) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    if (!isValidDate(date)) return [["Invalid date format."]];
    
    const url = `https://api.polygon.io/v2/aggs/grouped/locale/global/market/fx/${date}?adjusted=true&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const results = [["Ticker", "Open", "High", "Low", "Close"]];
    data.results.forEach(bar => {
      results.push([bar.T, bar.o, bar.h, bar.l, bar.c]);
    });
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves the last quote for a forex pair.
 * @customfunction
 * @param {string} pair The forex pair.
 * @returns {Promise<string[][]>} Array with last quote details.
 */
export async function getForexLastQuote(pair) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v2/last/nbbo/${pair}?apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const quote = data.results;
    return [
      ["Bid", "Ask", "Timestamp"],
      [quote.bp, quote.ap, new Date(quote.t).toISOString()]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves a snapshot of a forex pair.
 * @customfunction
 * @param {string} pair The forex pair.
 * @returns {Promise<string[][]>} Array with snapshot data.
 */
export async function getForexSnapshot(pair) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v3/snapshot?ticker.any=${pair}&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results || data.results.length === 0) return [["No data returned."]];
    
    const snap = data.results[0];
    return [
      ["Last Price", "Open", "High", "Low"],
      [snap.lastQuote.p, snap.day.o, snap.day.h, snap.day.l]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

// ---------------------------
// Crypto APIs
// ---------------------------

/**
 * Retrieves historical OHLC data for a crypto pair.
 * @customfunction
 * @param {string} pair The crypto pair (e.g., "X:BTCUSD").
 * @param {string} fromDate Start date in YYYY-MM-DD format.
 * @param {string} toDate End date in YYYY-MM-DD format.
 * @param {string} [timespan="day"] Timespan ("minute", "hour", "day").
 * @returns {Promise<string[][]>} Array of OHLC data.
 */
export async function getCryptoHistoricalOHLC(pair, fromDate, toDate, timespan = "day") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    if (!isValidDate(fromDate) || !isValidDate(toDate)) return [["Invalid date format."]];
    
    const url = `https://api.polygon.io/v2/aggs/ticker/${pair}/range/1/${timespan}/${fromDate}/${toDate}?adjusted=true&sort=asc&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const results = [["Date", "Open", "High", "Low", "Close", "Volume"]];
    data.results.forEach(bar => {
      const date = new Date(bar.t).toISOString().split('T')[0];
      results.push([date, bar.o, bar.h, bar.l, bar.c, bar.v]);
    });
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves grouped daily OHLC for all crypto pairs.
 * @customfunction
 * @param {string} date Date in YYYY-MM-DD format.
 * @returns {Promise<string[][]>} Array of daily OHLC data.
 */
export async function getCryptoGroupedDailyBars(date) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    if (!isValidDate(date)) return [["Invalid date format."]];
    
    const url = `https://api.polygon.io/v2/aggs/grouped/locale/global/market/crypto/${date}?adjusted=true&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const results = [["Ticker", "Open", "High", "Low", "Close", "Volume"]];
    data.results.forEach(bar => {
      results.push([bar.T, bar.o, bar.h, bar.l, bar.c, bar.v]);
    });
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves the last trade for a crypto pair.
 * @customfunction
 * @param {string} pair The crypto pair.
 * @returns {Promise<string[][]>} Array with last trade details.
 */
export async function getCryptoLastTrade(pair) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v2/last/trade/${pair}?apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const trade = data.results;
    return [
      ["Price", "Size", "Timestamp"],
      [trade.p, trade.s, new Date(trade.t).toISOString()]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves the last quote for a crypto pair.
 * @customfunction
 * @param {string} pair The crypto pair.
 * @returns {Promise<string[][]>} Array with last quote details.
 */
export async function getCryptoLastQuote(pair) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v2/last/nbbo/${pair}?apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const quote = data.results;
    return [
      ["Bid", "Ask", "Timestamp"],
      [quote.bp, quote.ap, new Date(quote.t).toISOString()]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves a snapshot of a crypto pair.
 * @customfunction
 * @param {string} pair The crypto pair.
 * @returns {Promise<string[][]>} Array with snapshot data.
 */
export async function getCryptoSnapshot(pair) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v3/snapshot?ticker.any=${pair}&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results || data.results.length === 0) return [["No data returned."]];
    
    const snap = data.results[0];
    return [
      ["Last Price", "Open", "High", "Low", "Volume"],
      [snap.lastTrade.p, snap.day.o, snap.day.h, snap.day.l, snap.day.v]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

// ---------------------------
// Reference APIs
// ---------------------------

/**
 * Retrieves a list of available markets.
 * @customfunction
 * @returns {Promise<string[][]>} Array of market information.
 */
export async function getMarkets() {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v3/reference/markets?apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const results = [["Market", "Description"]];
    data.results.forEach(market => {
      results.push([market.market, market.desc]);
    });
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves a list of available locales.
 * @customfunction
 * @returns {Promise<string[][]>} Array of locale information.
 */
export async function getLocales() {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v3/reference/locales?apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const results = [["Locale", "Name"]];
    data.results.forEach(locale => {
      results.push([locale.locale, locale.name]);
    });
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves company financials for a ticker.
 * @customfunction
 * @param {string} ticker The stock ticker (e.g., "AAPL").
 * @param {string} [period="annual"] Period ("annual" or "quarterly").
 * @param {number} [limit=1] Number of statements to retrieve.
 * @returns {Promise<string[][]>} Array of financial data.
 * @note This function may require a paid Polygon.io plan.
 */
export async function getCompanyFinancials(ticker, period = "annual", limit = 1) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/vX/reference/financials?ticker=${ticker}&timeframe=${period}&limit=${limit}&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const financials = data.results[0].financials.income_statement;
    return [
      ["Metric", "Value"],
      ["Revenues", financials.revenues?.value || "N/A"],
      ["Net Income", financials.net_income?.value || "N/A"],
      ["Operating Expenses", financials.operating_expenses?.value || "N/A"]
    ];
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves a list of trade conditions.
 * @customfunction
 * @param {string} [assetClass="stocks"] Asset class ("stocks", "options", "crypto", "fx").
 * @returns {Promise<string[][]>} Array of condition information.
 */
export async function getConditions(assetClass = "stocks") {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v3/reference/conditions?asset_class=${assetClass}&apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const results = [["ID", "Name", "Type"]];
    data.results.forEach(condition => {
      results.push([condition.id, condition.name, condition.type]);
    });
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Retrieves a list of ticker types.
 * @customfunction
 * @returns {Promise<string[][]>} Array of ticker type information.
 */
export async function getTickerTypes() {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return [["API key not set."]];
    
    const url = `https://api.polygon.io/v3/reference/ticker-types?apiKey=${apiKey}`;
    const response = await fetch(url);
    if (!response.ok) return [[`HTTP error! status: ${response.status}`]];
    
    const data = await response.json();
    if (!data.results) return [["No data returned."]];
    
    const results = [["Code", "Description", "Asset Class"]];
    data.results.forEach(type => {
      results.push([type.code, type.description, type.asset_class]);
    });
    return results;
  } catch (error) {
    return [[`Error: ${error.message}`]];
  }
}