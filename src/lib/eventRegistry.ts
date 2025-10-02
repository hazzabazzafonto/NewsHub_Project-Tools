import fetch from 'node-fetch';

// Event Registry configuration is now handled via environment variables

export interface EventRegistryArticle {
  id: string;
  title: string;
  url?: string;
  source?: {
    title: string;
    uri: string;
  };
  authors?: Array<{
    name: string;
    uri?: string;
  }>;
  body?: string;
  summary?: string;
  dateTime?: string;
  date?: string;
}

export interface EventRegistryResponse {
  articles: {
    results: EventRegistryArticle[];
    totalResults: number;
  };
}

export interface EventRegistryRequest {
  query: {
    $query: {
      $and: any[];
    };
    $filter?: {
      dataType?: string[];
    };
  };
  resultType: string;
  articlesSortBy?: string;
  articlesPage?: number;
  articlesCount?: number;
  apiKey: string;
}

export class EventRegistryAPI {
  private apiUrl = 'https://eventregistry.org/api/v1/article/getArticles';
  private articlesPerPage: number;
  private requestDelayMs: number;
  private maxRetries: number;
  private timeoutMs: number;

  constructor() {
    // Configuration is now handled via environment variables
    this.articlesPerPage = parseInt(process.env.EVENT_REGISTRY_ARTICLES_PER_PAGE || '100');
    this.requestDelayMs = parseInt(process.env.EVENT_REGISTRY_REQUEST_DELAY_MS || '1000');
    this.maxRetries = parseInt(process.env.EVENT_REGISTRY_MAX_RETRIES || '3');
    this.timeoutMs = parseInt(process.env.EVENT_REGISTRY_TIMEOUT_MS || '30000');
  }

  /**
   * Fetch articles from Event Registry API with pagination
   */
  async fetchArticles(requestBody: EventRegistryRequest): Promise<EventRegistryArticle[]> {
    const allArticles: EventRegistryArticle[] = [];
    let currentPage = 1;
    let hasMorePages = true;

    while (hasMorePages) {
      const request = {
        ...requestBody,
        articlesPage: currentPage,
        articlesCount: this.articlesPerPage,
        apiKey: process.env.EVENT_REGISTRY_API_KEY || ''
      };

      console.log(`Fetching page ${currentPage}...`);

      try {
        const response = await fetch(this.apiUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(request)
        });

        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data: EventRegistryResponse = await response.json();
        const articles = data.articles?.results || [];

        allArticles.push(...articles);

        console.log(`Fetched ${articles.length} articles from page ${currentPage}`);
        console.log(`Total articles so far: ${allArticles.length}`);

        // Check if there are more pages
        hasMorePages = articles.length === this.articlesPerPage && 
                      allArticles.length < (data.articles?.totalResults || 0);

        if (hasMorePages) {
          currentPage++;
          // Small delay to avoid rate limiting
          await new Promise(resolve => setTimeout(resolve, this.requestDelayMs));
        }
      } catch (error) {
        console.error(`Error fetching page ${currentPage}:`, error);
        throw error;
      }
    }

    return allArticles;
  }

  /**
   * Build a request body for Event Registry API
   */
  buildRequest(params: {
    searchTerms: string[];
    sources: string[];
    startDate: string;
    endDate: string;
    useBooleanQuery?: boolean;
    booleanQuery?: string;
  }): EventRegistryRequest {
    const { searchTerms, sources, startDate, endDate, useBooleanQuery, booleanQuery } = params;

    const query: EventRegistryRequest['query'] = {
      $query: {
        $and: []
      },
      $filter: {
        dataType: ["news", "blog"]
      }
    };

    // Add search terms
    if (useBooleanQuery && booleanQuery) {
      // Parse boolean query and convert to Event Registry format
      const parsedQuery = this.parseBooleanQuery(booleanQuery);
      if (parsedQuery) {
        query.$query.$and.push(parsedQuery);
      }
    } else {
      // Simple OR logic for multiple terms
      if (searchTerms.length === 1) {
        query.$query.$and.push({
          keyword: searchTerms[0],
          keywordLoc: "body"
        });
      } else if (searchTerms.length > 1) {
        const termConditions = searchTerms.map(term => ({
          keyword: term,
          keywordLoc: "body"
        }));
        query.$query.$and.push({
          $or: termConditions
        });
      }
    }

    // Add sources filter
    if (sources.length > 0) {
      const sourceConditions = sources.map(source => ({
        sourceUri: source
      }));
      query.$query.$and.push({
        $or: sourceConditions
      });
    }

    // Add date range
    query.$query.$and.push({
      dateStart: startDate,
      dateEnd: endDate
    });

    const request = {
      query,
      resultType: "articles",
      articlesSortBy: "date",
      apiKey: process.env.EVENT_REGISTRY_API_KEY || ''
    };

    // Log the request for debugging
    console.log('Event Registry Request:', JSON.stringify(request, null, 2));

    return request;
  }

  /**
   * Format articles for Google Sheets
   */
  formatArticlesForSheets(articles: EventRegistryArticle[]): string[][] {
    return articles.map((article, index) => {
      // Extract source name
      const sourceName = article.source?.title || 'Unknown Source';
      
      // Extract title
      const title = article.title || 'No Title';
      
      // Extract authors
      const authors = this.extractAuthors(article);
      
      // Extract URL
      const url = article.url || '';
      
      // Extract content (body or summary)
      const content = article.body || article.summary || 'No content available';
      
      // Extract date
      const date = article.dateTime || article.date || new Date().toISOString().split('T')[0];
      
      // Article ID (simple row number)
      const articleId = (index + 1).toString();
      
      // Input method
      const inputMethod = 'Event Registry';

      return [articleId, sourceName, title, authors, url, content, date, inputMethod];
    });
  }

  /**
   * Parse boolean query string into Event Registry format
   * Supports complex queries with nested AND/OR operations and parentheses
   */
  private parseBooleanQuery(queryString: string): any {
    try {
      // Remove quotes and clean up the query
      const cleanQuery = queryString.replace(/"/g, '').trim();
      
      // Parse the query recursively to handle nested operations
      return this.parseQueryExpression(cleanQuery);
    } catch (error) {
      console.error('Error parsing boolean query:', error);
      // Fallback to simple keyword search
      return {
        keyword: queryString,
        keywordLoc: "body"
      };
    }
  }

  /**
   * Recursively parse query expressions with proper operator precedence
   */
  private parseQueryExpression(expression: string): any {
    // Remove outer parentheses if they wrap the entire expression
    expression = expression.trim();
    if (expression.startsWith('(') && expression.endsWith(')')) {
      const inner = expression.slice(1, -1);
      if (this.isBalanced(inner)) {
        expression = inner;
      }
    }

    // Handle OR operations (lower precedence)
    const orIndex = this.findOperatorIndex(expression, ' OR ');
    if (orIndex !== -1) {
      const left = expression.substring(0, orIndex).trim();
      const right = expression.substring(orIndex + 4).trim();
      
      return {
        $or: [
          this.parseQueryExpression(left),
          this.parseQueryExpression(right)
        ]
      };
    }

    // Handle AND operations (higher precedence)
    const andIndex = this.findOperatorIndex(expression, ' AND ');
    if (andIndex !== -1) {
      const left = expression.substring(0, andIndex).trim();
      const right = expression.substring(andIndex + 5).trim();
      
      return {
        $and: [
          this.parseQueryExpression(left),
          this.parseQueryExpression(right)
        ]
      };
    }

    // Base case: single term
    return {
      keyword: expression,
      keywordLoc: "body"
    };
  }

  /**
   * Find the index of an operator, respecting parentheses
   */
  private findOperatorIndex(expression: string, operator: string): number {
    let depth = 0;
    for (let i = 0; i < expression.length - operator.length + 1; i++) {
      if (expression[i] === '(') depth++;
      else if (expression[i] === ')') depth--;
      else if (depth === 0 && expression.substring(i, i + operator.length) === operator) {
        return i;
      }
    }
    return -1;
  }

  /**
   * Check if parentheses are balanced
   */
  private isBalanced(expression: string): boolean {
    let count = 0;
    for (const char of expression) {
      if (char === '(') count++;
      else if (char === ')') count--;
      if (count < 0) return false;
    }
    return count === 0;
  }

  /**
   * Extract authors from article data
   */
  private extractAuthors(article: EventRegistryArticle): string {
    if (article.authors && Array.isArray(article.authors) && article.authors.length > 0) {
      const authorNames = article.authors
        .map(author => author.name || author.uri || 'Unknown Author')
        .filter(name => name && name.trim().length > 0);
      
      return authorNames.length > 0 ? authorNames.join(', ') : 'No Author Available';
    }
    
    return 'No Author Available';
  }
}
