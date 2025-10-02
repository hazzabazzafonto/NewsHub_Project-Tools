import express from 'express';
import { EventRegistryAPI } from '../lib/eventRegistry';
import { getAuthClient } from '../lib/googleSheets';
import { google } from 'googleapis';

/**
 * Get the next sequential article ID from the sheet
 */
async function getNextArticleId(sheetId: string): Promise<number> {
  try {
    const authClient = await getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    
    // Read the first column to get existing IDs
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: 'Articles!A:A',
    });
    
    const rows = response.data.values || [];
    
    if (rows.length <= 1) {
      // Only headers or no data, start from 1
      return 1;
    }
    
    // Skip header row and find the highest ID
    const ids = rows.slice(1)
      .map(row => parseInt(row[0]) || 0)
      .filter(id => !isNaN(id));
    
    return ids.length > 0 ? Math.max(...ids) + 1 : 1;
  } catch (error) {
    console.error('Error getting next article ID:', error);
    return 1; // Fallback to 1
  }
}

const router = express.Router();

// Initialize Event Registry API
const eventRegistryAPI = new EventRegistryAPI();

/**
 * Get sources from Google Sheets
 */
router.get('/sources', async (req: express.Request, res: express.Response) => {
  try {
    const sourcesSheetId = process.env.EVENT_REGISTRY_SOURCES_SHEET_ID;
    const sourcesRange = process.env.EVENT_REGISTRY_SOURCES_RANGE || 'Sheet1!A2:E';
    
    if (!sourcesSheetId) {
      return res.status(400).json({ 
        message: 'Sources sheet ID not configured. Please set EVENT_REGISTRY_SOURCES_SHEET_ID environment variable.' 
      });
    }

    const authClient = await getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    const sourcesData = await sheets.spreadsheets.values.get({
      spreadsheetId: sourcesSheetId,
      range: sourcesRange,
    });

    const sources = [];
    const regions = new Set();
    const countries = new Set();
    const languages = new Set();

    if (sourcesData.data.values) {
      for (const row of sourcesData.data.values) {
        const title = row[0]?.trim() || '';
        const region = row[1]?.trim() || '';
        const country = row[2]?.trim() || '';
        const language = row[3]?.trim() || '';
        const uri = row[4]?.trim() || '';

        if (title && uri) {
          sources.push({
            title,
            region,
            country,
            language,
            uri: cleanSourceUrl(uri),
            selected: false
          });

          if (region) regions.add(region);
          if (country) countries.add(country);
          if (language) languages.add(language);
        }
      }
    }

    res.json({
      sources,
      filters: {
        regions: Array.from(regions).sort(),
        countries: Array.from(countries).sort(),
        languages: Array.from(languages).sort()
      }
    });
  } catch (error: any) {
    console.error('Error fetching sources:', error);
    res.status(500).json({ message: error.message || 'Failed to fetch sources' });
  }
});

/**
 * Fetch articles from Event Registry
 */
router.post('/fetch-articles', async (req: express.Request, res: express.Response) => {
  try {
    const { 
      searchTerms, 
      sources, 
      startDate, 
      endDate, 
      useBooleanQuery, 
      booleanQuery,
      projectId 
    } = req.body;

    // Validate required fields
    if (!startDate || !endDate) {
      return res.status(400).json({ message: 'Start date and end date are required' });
    }

    if (!searchTerms || searchTerms.length === 0) {
      return res.status(400).json({ message: 'At least one search term is required' });
    }

    if (!sources || sources.length === 0) {
      return res.status(400).json({ message: 'At least one source is required' });
    }

    // Build request
    const requestBody = eventRegistryAPI.buildRequest({
      searchTerms,
      sources,
      startDate,
      endDate,
      useBooleanQuery,
      booleanQuery
    });

    // Fetch articles
    const articles = await eventRegistryAPI.fetchArticles(requestBody);

    // Format articles for sheets
    const formattedArticles = eventRegistryAPI.formatArticlesForSheets(articles);

    // Add headers
    const headers = [
      'Article ID', 
      'Article Source Outlet', 
      'Article Title', 
      'Article Author/s', 
      'Article URLs', 
      'Article Full Body Text', 
      'Date the article was written', 
      'Article Input Method'
    ];

    const sheetData = [headers, ...formattedArticles];

    res.json({
      success: true,
      articles: articles,
      formattedData: sheetData,
      count: articles.length,
      message: `Successfully fetched ${articles.length} articles`
    });

  } catch (error: any) {
    console.error('Error fetching articles:', error);
    res.status(500).json({ 
      message: error.message || 'Failed to fetch articles from Event Registry' 
    });
  }
});

/**
 * Write articles to project sheet
 */
router.post('/write-to-sheet', async (req: express.Request, res: express.Response) => {
  try {
    const { projectId, articles } = req.body;

    if (!projectId) {
      return res.status(400).json({ message: 'Project ID is required' });
    }

    if (!articles || articles.length === 0) {
      return res.status(400).json({ message: 'No articles to write' });
    }

    // Get project to find sheet ID
    const { getProject } = await import('../lib/db');
    const project = await getProject(projectId);
    
    if (!project) {
      return res.status(404).json({ message: 'Project not found' });
    }

    // Get fresh auth client to avoid token expiration issues
    const authClient = await getAuthClient();
    
    // Refresh token if needed
    try {
      await authClient.getAccessToken();
    } catch (error) {
      console.log('Refreshing authentication token...');
      await authClient.refreshAccessToken();
    }
    
    const sheets = google.sheets({ 
      version: 'v4', 
      auth: authClient,
      timeout: 60000 // 60 second timeout for the entire client
    });

    // Determine if these are Event Registry articles or manual entries
    const isEventRegistryArticles = articles.length > 0 && articles[0].source?.title !== undefined;
    
    let formattedArticles: string[][];
    
    // Get next sequential ID for all articles
    const nextId = await getNextArticleId(project.sheetId);
    
    if (isEventRegistryArticles) {
      // Event Registry articles
      formattedArticles = eventRegistryAPI.formatArticlesForSheets(articles, nextId);
    } else {
      // Manual entry articles
      formattedArticles = eventRegistryAPI.formatManualArticlesForSheets(articles, nextId);
    }

    const sheetData = formattedArticles;

    // Write to the project's sheet in smaller batches with retry logic
    const batchSize = 100; // Reduced batch size to avoid timeouts
    const totalBatches = Math.ceil(sheetData.length / batchSize);
    const maxRetries = 3;
    const retryDelay = 2000; // 2 seconds between retries
    
    console.log(`Writing ${sheetData.length} articles in ${totalBatches} batches of ${batchSize}`);
    
    for (let i = 0; i < totalBatches; i++) {
      const startIndex = i * batchSize;
      const endIndex = Math.min(startIndex + batchSize, sheetData.length);
      const batch = sheetData.slice(startIndex, endIndex);
      
      console.log(`Processing batch ${i + 1}/${totalBatches} (${batch.length} articles)`);
      
      // Retry logic for each batch
      let success = false;
      let attempt = 0;
      
      while (!success && attempt < maxRetries) {
        try {
          attempt++;
          console.log(`Batch ${i + 1} attempt ${attempt}/${maxRetries}`);
          
          await sheets.spreadsheets.values.append({
            spreadsheetId: project.sheetId,
            range: 'Articles!A:A',
            valueInputOption: 'RAW',
            requestBody: {
              values: batch,
            },
          }, {
            timeout: 60000 // 60 second timeout
          });
          
          success = true;
          console.log(`Batch ${i + 1} completed successfully`);
          
        } catch (error: any) {
          console.error(`Batch ${i + 1} attempt ${attempt} failed:`, error.message);
          
          if (attempt < maxRetries) {
            console.log(`Retrying batch ${i + 1} in ${retryDelay}ms...`);
            await new Promise(resolve => setTimeout(resolve, retryDelay));
          } else {
            console.error(`Batch ${i + 1} failed after ${maxRetries} attempts`);
            throw new Error(`Failed to write batch ${i + 1} after ${maxRetries} attempts: ${error.message}`);
          }
        }
      }
      
      // Add delay between batches to avoid rate limiting
      if (i < totalBatches - 1) {
        await new Promise(resolve => setTimeout(resolve, 1000)); // Increased to 1 second
      }
    }

    console.log(`Successfully wrote all ${articles.length} articles to project sheet`);
    
    res.json({
      success: true,
      message: `Successfully wrote ${articles.length} articles to project sheet${totalBatches > 1 ? ` in ${totalBatches} batches` : ''}`,
      count: articles.length,
      batches: totalBatches,
      batchSize: batchSize
    });

  } catch (error: any) {
    console.error('Error writing to sheet:', error);
    res.status(500).json({ 
      message: error.message || 'Failed to write articles to sheet' 
    });
  }
});

/**
 * Helper function to clean source URLs
 */
function cleanSourceUrl(source: string): string {
  return source
    .toLowerCase()
    .replace(/^https?:\/\//, '') // Remove http:// or https://
    .replace(/^www\./, '') // Remove www.
    .replace(/\/$/, '') // Remove trailing slash
    .trim();
}

export default router;
