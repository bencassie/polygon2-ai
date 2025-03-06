Office.onReady(() => {
  // Hide sideload message and show app body
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  // Define custom functions with signature and example formula
  const customFunctions = [
    {
      signature: "getAnalystRecommendations(ticker)",
      example: '=POLYGON.getAnalystRecommendations("AAPL")',
      description: "Retrieves analyst recommendations for a stock"
    },
    {
      signature: "getATR(ticker, [period])",
      example: '=POLYGON.getATR("AAPL", 14)',
      description: "Calculates Average True Range (ATR) for a stock"
    },
    {
      signature: "getBatchQuotes(tickers)",
      example: '=POLYGON.getBatchQuotes({"AAPL","MSFT","GOOGL"})',
      description: "Retrieves latest quotes for multiple tickers"
    },
    {
      signature: "getBollingerBands(ticker, [period], [stdDev])",
      example: '=POLYGON.getBollingerBands("AAPL", 20, 2)',
      description: "Calculates Bollinger Bands for a stock"
    },
    {
      signature: "getCorporateActions(ticker)",
      example: '=POLYGON.getCorporateActions("AAPL")',
      description: "Retrieves corporate actions for a stock"
    },
    {
      signature: "getCryptoData(cryptoPair)",
      example: '=POLYGON.getCryptoData("BTC-USD")',
      description: "Retrieves cryptocurrency market data"
    },
    {
      signature: "getDividends(ticker, [limit])",
      example: '=POLYGON.getDividends("AAPL", 4)',
      description: "Retrieves dividend information for a ticker"
    },
    {
      signature: "getEarningsAnalysis(ticker)",
      example: '=POLYGON.getEarningsAnalysis("AAPL")',
      description: "Retrieves earnings analysis data for a stock"
    },
    {
      signature: "getEconomicCalendar(fromDate, toDate)",
      example: '=POLYGON.getEconomicCalendar("2024-01-01", "2024-12-31")',
      description: "Retrieves economic events and indicators"
    },
    {
      signature: "getExchanges([exchangeType])",
      example: '=POLYGON.getExchanges("exchange")',
      description: "Retrieves information about stock exchanges"
    },
    {
      signature: "getFibonacciLevels(ticker, fromDate, toDate)",
      example: '=POLYGON.getFibonacciLevels("AAPL", "2024-01-01", "2024-03-31")',
      description: "Calculates Fibonacci retracement levels"
    },
    {
      signature: "getForexRates(fromCurrency, toCurrency)",
      example: '=POLYGON.getForexRates("USD", "EUR")',
      description: "Retrieves forex exchange rates"
    },
    {
      signature: "getGroupedDailyBars(date, [locale], [market])",
      example: '=POLYGON.getGroupedDailyBars("2024-03-21", "us", "stocks")',
      description: "Retrieves grouped daily bars for the entire market"
    },
    {
      signature: "getHistoricalOHLC(ticker, fromDate, toDate, [timespan])",
      example: '=POLYGON.getHistoricalOHLC("AAPL", "2023-01-01", "2023-12-31", "day")',
      description: "Retrieves historical OHLC data"
    },
    {
      signature: "getIndicesData(indices)",
      example: '=POLYGON.getIndicesData({"SPX","DJI"})',
      description: "Retrieves data for market indices"
    },
    {
      signature: "getInsiderTrading(ticker, [limit])",
      example: '=POLYGON.getInsiderTrading("AAPL", 20)',
      description: "Retrieves insider trading activity"
    },
    {
      signature: "getInstitutionalOwnership(ticker)",
      example: '=POLYGON.getInstitutionalOwnership("AAPL")',
      description: "Retrieves institutional ownership data"
    },
    {
      signature: "getLatestPrice(ticker, [property])",
      example: '=POLYGON.getLatestPrice("AAPL")',
      description: "Retrieves latest price data"
    },
    {
      signature: "getMarketBreadth([market])",
      example: '=POLYGON.getMarketBreadth("stocks")',
      description: "Retrieves market breadth indicators"
    },
    {
      signature: "getMarketDepth(ticker)",
      example: '=POLYGON.getMarketDepth("AAPL")',
      description: "Retrieves market depth data"
    },
    {
      signature: "getMarketHolidays([year])",
      example: '=POLYGON.getMarketHolidays(2024)',
      description: "Gets upcoming market holidays"
    },
    {
      signature: "getMarketStatus([market])",
      example: '=POLYGON.getMarketStatus("us")',
      description: "Gets current market status"
    },
    {
      signature: "getOptionsChain(ticker, expirationDate)",
      example: '=POLYGON.getOptionsChain("AAPL", "2024-06-21")',
      description: "Retrieves options chain data"
    },
    {
      signature: "getOptionsGreeks(optionTicker)",
      example: '=POLYGON.getOptionsGreeks("O:AAPL240621C00150000")',
      description: "Calculates options Greeks"
    },
    {
      signature: "getOptionsStrategy(ticker, strategy)",
      example: '=POLYGON.getOptionsStrategy("AAPL", "covered_call")',
      description: "Analyzes options strategies"
    },
    {
      signature: "getPivotPoints(ticker, [method])",
      example: '=POLYGON.getPivotPoints("AAPL", "standard")',
      description: "Calculates pivot points"
    },
    {
      signature: "getPortfolioSummary(portfolioData)",
      example: '=POLYGON.getPortfolioSummary(A1:B10)',
      description: "Creates portfolio summary"
    },
    {
      signature: "getPreviousClose(tickers)",
      example: '=POLYGON.getPreviousClose({"AAPL","MSFT"})',
      description: "Retrieves previous close data"
    },
    {
      signature: "getRiskIndicators(ticker)",
      example: '=POLYGON.getRiskIndicators("AAPL")',
      description: "Calculates risk metrics"
    },
    {
      signature: "screenStocks(minPrice, maxPrice, minVolume, minChange, maxChange, [limit])",
      example: '=POLYGON.screenStocks(10, 100, 1000000, -5, 5, 50)',
      description: "Screens stocks based on criteria"
    },
    {
      signature: "getSectorPerformance([timespan])",
      example: '=POLYGON.getSectorPerformance("week")',
      description: "Analyzes sector performance"
    },
    {
      signature: "getSocialSentiment(ticker, fromDate, toDate)",
      example: '=POLYGON.getSocialSentiment("AAPL", "2024-01-01", "2024-03-31")',
      description: "Retrieves social sentiment metrics"
    },
    {
      signature: "getStockCorrelation(ticker1, ticker2, [days])",
      example: '=POLYGON.getStockCorrelation("AAPL", "MSFT", 30)',
      description: "Calculates correlation between stocks"
    },
    {
      signature: "getStockSplits(ticker, [limit])",
      example: '=POLYGON.getStockSplits("AAPL", 5)',
      description: "Retrieves stock splits history"
    },
    {
      signature: "getTechnicalIndicator(ticker, indicator, [period], [from], [to])",
      example: '=POLYGON.getTechnicalIndicator("AAPL", "SMA", 14)',
      description: "Calculates technical indicators"
    },
    {
      signature: "getTickerDetails(ticker, [property])",
      example: '=POLYGON.getTickerDetails("AAPL")',
      description: "Gets ticker details"
    },
    {
      signature: "getTickerNews(ticker, [limit])",
      example: '=POLYGON.getTickerNews("AAPL", 5)',
      description: "Retrieves news articles"
    },
    {
      signature: "getTopMovers([type], [limit])",
      example: '=POLYGON.getTopMovers("gainers", 20)',
      description: "Gets top gaining/losing stocks"
    },
    {
      signature: "searchTickers(searchTerm, [limit])",
      example: '=POLYGON.searchTickers("Apple", 10)',
      description: "Searches for tickers"
    }
  ];

  
  // Populate the functions list
  const functionsList = document.getElementById("functions-list");
  customFunctions.forEach(func => {
      const functionDiv = document.createElement("div");
      functionDiv.className = "function";

      const signatureHeader = document.createElement("h3");
      signatureHeader.textContent = func.signature;
      functionDiv.appendChild(signatureHeader);

      const insertButton = document.createElement("button");
      insertButton.textContent = "Insert";
      insertButton.onclick = () => insertFunction(func.example);
      functionDiv.appendChild(insertButton);

      functionsList.appendChild(functionDiv);
  });

  // API key handling
  const savedKey = localStorage.getItem("polygonApiKey");
  toggleApiKeyInput(!savedKey);

  document.getElementById("setApiKey").onclick = setApiKey;
  document.getElementById("changeApiKey").onclick = changeApiKey;
});

// Function to insert a formula into the active cell
async function insertFunction(formula) {
  try {
      await Excel.run(async (context) => {
          const activeCell = context.workbook.getActiveCell();
          activeCell.formulas = [[formula]];
          await context.sync();
      });
  } catch (error) {
      console.error("Error inserting formula:", error);
  }
}

// Function to toggle API key input visibility
function toggleApiKeyInput(showInput) {
  const inputContainer = document.getElementById("apiKeyInputContainer");
  const apiKeySet = document.getElementById("apiKeySet");
  if (showInput) {
      inputContainer.style.display = "block";
      apiKeySet.style.display = "none";
  } else {
      inputContainer.style.display = "none";
      apiKeySet.style.display = "block";
  }
}

// Function to save the API key
function setApiKey() {
  const input = document.getElementById("apiKeyInput").value;
  localStorage.setItem("polygonApiKey", input);
  document.getElementById("apiKeyStatus").textContent = "API key saved successfully.";
  toggleApiKeyInput(!input);
}

// Function to change the API key
function changeApiKey() {
  document.getElementById("apiKeyInput").value = "";
  toggleApiKeyInput(true);
}