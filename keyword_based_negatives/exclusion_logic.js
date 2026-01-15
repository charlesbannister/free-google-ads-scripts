/**
 * Exclusion Logic Helper Script
 * Run with Node.js to test whether a search term should be excluded
 * 
 * Usage: node exclusion_logic.js
 * 
 * @authors Charles Bannister & Gabriele Benedetti
 */

// ============================================================================
// CONFIGURATION - Update these to test different scenarios
// ============================================================================

/**
 * Fuzzy match threshold (0-100)
 * Higher = stricter matching (fewer exclusions)
 * Lower = more lenient matching (more exclusions)
 */
const FUZZY_MATCH_THRESHOLD = 80;

/**
 * Test cases - each object contains keywords, a search term, and expected result
 * shouldExclude: true = search term should be excluded (added as negative)
 * shouldExclude: false = search term should be kept (matches keywords)
 */
const TEST_CASES = [
  // Exact and near-exact matches
  { searchTerm: 'running shoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'nike running shoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // No space variations
  { searchTerm: 'runningshoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'runningshoe', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'r u n n i n g s h o e', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // Singular vs plural
  { searchTerm: 'running shoe', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'shoe running', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // Misspellings - running
  { searchTerm: 'runing shoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'runnign shoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'runnnig shoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'runnin shoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // Misspellings - shoes
  { searchTerm: 'running sheos', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'running shoess', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'running shose', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // Combined misspellings - now matches via shared word check (shoes ~ sheos)
  { searchTerm: 'runing sheos', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'runnig shoe', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // Extra words (should still match)
  { searchTerm: 'best running shoes 2024', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'cheap running shoes online', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // Subset of keyword words (search term words exist in longer keyword)
  { searchTerm: 'osmosis water filter', keywords: ['reverse osmosis water filter'], shouldExclude: false },
  { searchTerm: 'running shoes', keywords: ['best running shoes', 'nike running shoes'], shouldExclude: false },

  // Shared word match (any keyword word exists in search term)
  { searchTerm: 'reverse osmosis ro', keywords: ['ro system'], shouldExclude: false },
  { searchTerm: 'dental autoclave repair', keywords: ['autoclave servicing'], shouldExclude: false },

  // Completely different (should exclude)
  { searchTerm: 'hiking boots', keywords: ['running shoes', 'nike running shoes'], shouldExclude: true },
  { searchTerm: 'basketball sneakers', keywords: ['running shoes', 'nike running shoes'], shouldExclude: true },
];

// ============================================================================
// MAIN EXECUTION
// ============================================================================

function runTests() {
  console.log('=== Exclusion Logic Tests ===');
  console.log(`Threshold: ${FUZZY_MATCH_THRESHOLD}%\n`);

  let passed = 0;
  let failed = 0;

  for (const testCase of TEST_CASES) {
    const result = shouldExcludeSearchTerm(testCase.searchTerm, testCase.keywords, FUZZY_MATCH_THRESHOLD);
    const actual = result.shouldExclude;
    const expected = testCase.shouldExclude;
    const testPassed = actual === expected;

    if (testPassed) {
      passed++;
    } else {
      failed++;
    }

    const passIcon = testPassed ? '✓' : '✗';
    const status = result.shouldExclude ? 'EXCLUDE' : 'KEEP';

    console.log(`${passIcon} ${status} | "${testCase.searchTerm}"`);
    console.log(`         Keywords: ${testCase.keywords.join(', ')}`);
    console.log(`         Score: ${result.bestMatchScore.toFixed(1)}% | Matched: ${result.bestMatchingKeyword || 'None'}`);

    if (!testPassed) {
      console.log(`         ⚠️  EXPECTED: ${expected ? 'EXCLUDE' : 'KEEP'}, GOT: ${actual ? 'EXCLUDE' : 'KEEP'}`);
    }

    console.log('');
  }

  console.log('=== Summary ===');
  console.log(`Passed: ${passed} | Failed: ${failed} | Total: ${TEST_CASES.length}`);
}

// ============================================================================
// FACADE FUNCTION - Main entry point for the logic
// ============================================================================

/**
 * Calculates the best fuzzy match score using a sliding window approach.
 * Checks if the keyword exists as a substring within the search term.
 * 
 * @param {string} keyword - The keyword to search for
 * @param {string} searchTerm - The search term to search within
 * @returns {number} Best match score from 0 to 100
 */
function getSlidingWindowScore(keyword, searchTerm) {
  let bestScore = 0;
  let startPosition = 0;
  let endPosition = keyword.length;

  while (endPosition <= searchTerm.length) {
    const windowScore = fuzzyMatchScore(
      keyword,
      searchTerm.substring(startPosition, endPosition)
    );
    if (windowScore > bestScore) {
      bestScore = windowScore;
    }
    startPosition++;
    endPosition++;
  }

  return bestScore;
}

/**
 * Determines if a search term should be excluded (added as negative).
 * A search term should be excluded if it doesn't match any keyword above the threshold.
 * 
 * @param {string} searchTerm - The search term to evaluate
 * @param {Array<string>} keywords - Array of keyword strings from the ad group
 * @param {number} threshold - Fuzzy match threshold (0-100), default 80
 * @returns {Object} Result object with shouldExclude, bestMatchScore, bestMatchingKeyword
 */
function shouldExcludeSearchTerm(searchTerm, keywords, threshold = 80) {
  if (!searchTerm || !keywords || keywords.length === 0) {
    return {
      shouldExclude: true,
      bestMatchScore: 0,
      bestMatchingKeyword: null
    };
  }

  const matcher = new SearchTermMatcher(false);
  let bestScore = 0;
  let bestKeyword = null;
  let foundMatch = false;

  for (const keyword of keywords) {
    const condition = {
      text: keyword,
      matchType: 'approx-contains',
      threshold: threshold
    };

    const isMatch = matcher.matchesCondition(searchTerm, condition);

    // Check multiple variations, use the highest score
    const normalScore = fuzzyMatchScore(keyword, searchTerm);
    const spacelessScore = fuzzyMatchScore(
      keyword.replace(/\s+/g, ''),
      searchTerm.replace(/\s+/g, '')
    );
    // Sort words alphabetically to ignore word order
    const sortedKeyword = keyword.toLowerCase().split(/\s+/).sort().join(' ');
    const sortedSearchTerm = searchTerm.toLowerCase().split(/\s+/).sort().join(' ');
    const sortedScore = fuzzyMatchScore(sortedKeyword, sortedSearchTerm);

    // Sliding window - check if keyword exists as substring in search term
    const slidingWindowScore = getSlidingWindowScore(keyword, searchTerm);

    // Subset words check - if all search term words exist in keyword
    const keywordWords = keyword.toLowerCase().split(/\s+/);
    const searchTermWords = searchTerm.toLowerCase().split(/\s+/);
    let subsetWordsScore = 0;
    const allWordsMatch = searchTermWords.every(searchWord =>
      keywordWords.some(keywordWord => fuzzyMatchScore(searchWord, keywordWord) >= threshold)
    );
    if (allWordsMatch) {
      const wordScores = searchTermWords.map(searchWord => {
        const scores = keywordWords.map(keywordWord => fuzzyMatchScore(searchWord, keywordWord));
        return Math.max(...scores);
      });
      subsetWordsScore = wordScores.reduce((a, b) => a + b, 0) / wordScores.length;
    }

    // Shared word check - if any keyword word matches any search term word
    let sharedWordScore = 0;
    for (const keywordWord of keywordWords) {
      for (const searchWord of searchTermWords) {
        const wordScore = fuzzyMatchScore(keywordWord, searchWord);
        sharedWordScore = Math.max(sharedWordScore, wordScore);
      }
    }

    // Individual words check - if any search term word matches the whole keyword
    let individualWordScore = 0;
    for (const searchWord of searchTermWords) {
      const wordScore = fuzzyMatchScore(keyword, searchWord);
      individualWordScore = Math.max(individualWordScore, wordScore);
    }

    const score = Math.max(normalScore, spacelessScore, sortedScore, slidingWindowScore, subsetWordsScore, sharedWordScore, individualWordScore);

    if (score > bestScore) {
      bestScore = score;
      bestKeyword = keyword;
    }

    if (isMatch) {
      foundMatch = true;
    }
  }

  return {
    shouldExclude: !foundMatch,
    bestMatchScore: bestScore,
    bestMatchingKeyword: bestKeyword
  };
}

// ============================================================================
// SEARCH TERM MATCHER CLASS
// ============================================================================

/**
 * Class for matching search terms against keywords using various matching methods.
 */
class SearchTermMatcher {
  /**
   * Creates a new SearchTermMatcher instance.
   * @param {boolean} debug - Enable debug logging
   */
  constructor(debug = false) {
    this.debug = debug;
  }

  /**
   * Checks if a search term matches a condition.
   * @param {string} term - The search term to check
   * @param {Object} condition - Condition object with text, matchType, and optional threshold
   * @returns {boolean} True if the term matches the condition
   */
  matchesCondition(term, condition) {
    if (!condition || !condition.text) {
      if (this.debug) console.warn('[SearchTermMatcher] Invalid condition:', condition);
      return false;
    }

    const termLower = term.toLowerCase();
    const textLower = condition.text.toLowerCase();

    switch (condition.matchType) {
      case 'contains':
        return termLower.includes(textLower);

      case 'not-contains':
        return !termLower.includes(textLower);

      case 'regex-contains':
        try {
          const regex = new RegExp(condition.text, 'i');
          return regex.test(term);
        } catch (error) {
          if (this.debug) console.error('[SearchTermMatcher] Invalid regex:', error);
          return false;
        }

      case 'approx-contains':
        return this.approxMatchContains(term, condition);

      default:
        return false;
    }
  }

  /**
   * Performs approximate (fuzzy) matching.
   * @param {string} term - The search term
   * @param {Object} condition - Condition with text and threshold
   * @returns {boolean} True if fuzzy match score exceeds threshold
   */
  approxMatchContains(term, condition) {
    const threshold = condition.threshold || 80;

    // Check whole term score
    const wholeTermScore = fuzzyMatchScore(condition.text, term);
    if (wholeTermScore >= threshold) {
      return true;
    }

    // Check without spaces
    const spacelessScore = fuzzyMatchScore(
      condition.text.replace(/\s+/g, ''),
      term.replace(/\s+/g, '')
    );
    if (spacelessScore >= threshold) {
      return true;
    }

    // Check with words sorted (ignore word order)
    const sortedKeyword = condition.text.toLowerCase().split(/\s+/).sort().join(' ');
    const sortedTerm = term.toLowerCase().split(/\s+/).sort().join(' ');
    const sortedScore = fuzzyMatchScore(sortedKeyword, sortedTerm);
    if (sortedScore >= threshold) {
      return true;
    }

    // Check if all search term words exist in keyword (subset match)
    // e.g. "osmosis water filter" should match "reverse osmosis water filter"
    const keywordWords = condition.text.toLowerCase().split(/\s+/);
    const searchTermWords = term.toLowerCase().split(/\s+/);
    const allWordsInKeyword = searchTermWords.every(searchWord =>
      keywordWords.some(keywordWord => fuzzyMatchScore(searchWord, keywordWord) >= threshold)
    );
    if (allWordsInKeyword) {
      return true;
    }

    // Check if any keyword word exists in search term words (shared word match)
    // e.g. "reverse osmosis ro" should match "ro system" because "ro" is shared
    const anySharedWord = keywordWords.some(keywordWord =>
      searchTermWords.some(searchWord => fuzzyMatchScore(keywordWord, searchWord) >= threshold)
    );
    if (anySharedWord) {
      return true;
    }

    // Check individual words
    const termWords = term.toLowerCase().split(' ');
    const anyWordMatches = termWords.some(
      word => fuzzyMatchScore(condition.text, word) >= threshold
    );
    if (anyWordMatches) {
      return true;
    }

    // Sliding window match
    let startPosition = 0;
    let endPosition = condition.text.length;
    while (endPosition <= term.length) {
      const windowScore = fuzzyMatchScore(
        condition.text,
        term.substring(startPosition, endPosition)
      );
      if (windowScore >= threshold) {
        return true;
      }
      startPosition++;
      endPosition++;
    }

    return false;
  }
}

// ============================================================================
// FUZZY SET IMPLEMENTATION
// ============================================================================

/**
 * FuzzySet implementation for fuzzy string matching.
 * Based on the FuzzySet.js library.
 */
var FuzzySet = (function () {
  "use strict";

  const FuzzySet = function (arr, useLevenshtein, gramSizeLower, gramSizeUpper) {
    var fuzzyset = {};

    arr = arr || [];
    fuzzyset.gramSizeLower = gramSizeLower || 2;
    fuzzyset.gramSizeUpper = gramSizeUpper || 3;
    fuzzyset.useLevenshtein = typeof useLevenshtein !== "boolean" ? true : useLevenshtein;

    fuzzyset.exactSet = {};
    fuzzyset.matchDict = {};
    fuzzyset.items = {};

    var levenshtein = function (str1, str2) {
      var current = [], prev, value;

      for (var i = 0; i <= str2.length; i++) {
        for (var j = 0; j <= str1.length; j++) {
          if (i && j) {
            if (str1.charAt(j - 1) === str2.charAt(i - 1)) {
              value = prev;
            } else {
              value = Math.min(current[j], current[j - 1], prev) + 1;
            }
          } else {
            value = i + j;
          }
          prev = current[j];
          current[j] = value;
        }
      }
      return current.pop();
    };

    var _distance = function (str1, str2) {
      if (str1 === null && str2 === null) throw "Trying to compare two null values";
      if (str1 === null || str2 === null) return 0;
      str1 = String(str1);
      str2 = String(str2);

      var distance = levenshtein(str1, str2);
      if (str1.length > str2.length) {
        return 1 - distance / str1.length;
      } else {
        return 1 - distance / str2.length;
      }
    };

    var _nonWordRe = /[^a-zA-Z0-9\u00C0-\u00FF\u0621-\u064A\u0660-\u0669, ]+/g;

    var _iterateGrams = function (value, gramSize) {
      gramSize = gramSize || 2;
      var simplified = "-" + value.toLowerCase().replace(_nonWordRe, "") + "-",
        lenDiff = gramSize - simplified.length,
        results = [];
      if (lenDiff > 0) {
        for (var i = 0; i < lenDiff; ++i) {
          simplified += "-";
        }
      }
      for (var i = 0; i < simplified.length - gramSize + 1; ++i) {
        results.push(simplified.slice(i, i + gramSize));
      }
      return results;
    };

    var _gramCounter = function (value, gramSize) {
      gramSize = gramSize || 2;
      var result = {},
        grams = _iterateGrams(value, gramSize),
        i = 0;
      for (i; i < grams.length; ++i) {
        if (grams[i] in result) {
          result[grams[i]] += 1;
        } else {
          result[grams[i]] = 1;
        }
      }
      return result;
    };

    fuzzyset.get = function (value, defaultValue, minMatchScore) {
      if (minMatchScore === undefined) {
        minMatchScore = 0.33;
      }
      var result = this._get(value, minMatchScore);
      if (!result && typeof defaultValue !== "undefined") {
        return defaultValue;
      }
      return result;
    };

    fuzzyset._get = function (value, minMatchScore) {
      var results = [];
      for (var gramSize = this.gramSizeUpper; gramSize >= this.gramSizeLower; --gramSize) {
        results = this.__get(value, gramSize, minMatchScore);
        if (results && results.length > 0) {
          return results;
        }
      }
      return null;
    };

    fuzzyset.__get = function (value, gramSize, minMatchScore) {
      var normalizedValue = this._normalizeStr(value),
        matches = {},
        gramCounts = _gramCounter(normalizedValue, gramSize),
        items = this.items[gramSize],
        sumOfSquareGramCounts = 0,
        gram, gramCount, i, index, otherGramCount;

      for (gram in gramCounts) {
        gramCount = gramCounts[gram];
        sumOfSquareGramCounts += Math.pow(gramCount, 2);
        if (gram in this.matchDict) {
          for (i = 0; i < this.matchDict[gram].length; ++i) {
            index = this.matchDict[gram][i][0];
            otherGramCount = this.matchDict[gram][i][1];
            if (index in matches) {
              matches[index] += gramCount * otherGramCount;
            } else {
              matches[index] = gramCount * otherGramCount;
            }
          }
        }
      }

      function isEmptyObject(obj) {
        for (var prop in obj) {
          if (obj.hasOwnProperty(prop)) return false;
        }
        return true;
      }

      if (isEmptyObject(matches)) {
        return null;
      }

      var vectorNormal = Math.sqrt(sumOfSquareGramCounts),
        results = [],
        matchScore;
      for (var matchIndex in matches) {
        matchScore = matches[matchIndex];
        results.push([
          matchScore / (vectorNormal * items[matchIndex][0]),
          items[matchIndex][1]
        ]);
      }
      var sortDescending = function (a, b) {
        if (a[0] < b[0]) return 1;
        if (a[0] > b[0]) return -1;
        return 0;
      };
      results.sort(sortDescending);
      if (this.useLevenshtein) {
        var newResults = [],
          endIndex = Math.min(50, results.length);
        for (var i = 0; i < endIndex; ++i) {
          newResults.push([_distance(results[i][1], normalizedValue), results[i][1]]);
        }
        results = newResults;
        results.sort(sortDescending);
      }
      newResults = [];
      results.forEach(function (scoreWordPair) {
        if (scoreWordPair[0] >= minMatchScore) {
          newResults.push([scoreWordPair[0], this.exactSet[scoreWordPair[1]]]);
        }
      }.bind(this));
      return newResults;
    };

    fuzzyset.add = function (value) {
      var normalizedValue = this._normalizeStr(value);
      if (normalizedValue in this.exactSet) {
        return false;
      }

      var i = this.gramSizeLower;
      for (i; i < this.gramSizeUpper + 1; ++i) {
        this._add(value, i);
      }
    };

    fuzzyset._add = function (value, gramSize) {
      var normalizedValue = this._normalizeStr(value),
        items = this.items[gramSize] || [],
        index = items.length;

      items.push(0);
      var gramCounts = _gramCounter(normalizedValue, gramSize),
        sumOfSquareGramCounts = 0,
        gram, gramCount;
      for (gram in gramCounts) {
        gramCount = gramCounts[gram];
        sumOfSquareGramCounts += Math.pow(gramCount, 2);
        if (gram in this.matchDict) {
          this.matchDict[gram].push([index, gramCount]);
        } else {
          this.matchDict[gram] = [[index, gramCount]];
        }
      }
      var vectorNormal = Math.sqrt(sumOfSquareGramCounts);
      items[index] = [vectorNormal, normalizedValue];
      this.items[gramSize] = items;
      this.exactSet[normalizedValue] = value;
    };

    fuzzyset._normalizeStr = function (str) {
      if (Object.prototype.toString.call(str) !== "[object String]") {
        throw "Must use a string as argument to FuzzySet functions";
      }
      return str.toLowerCase();
    };

    fuzzyset.length = function () {
      var count = 0, prop;
      for (prop in this.exactSet) {
        if (this.exactSet.hasOwnProperty(prop)) {
          count += 1;
        }
      }
      return count;
    };

    fuzzyset.isEmpty = function () {
      for (var prop in this.exactSet) {
        if (this.exactSet.hasOwnProperty(prop)) {
          return false;
        }
      }
      return true;
    };

    fuzzyset.values = function () {
      var values = [], prop;
      for (prop in this.exactSet) {
        if (this.exactSet.hasOwnProperty(prop)) {
          values.push(this.exactSet[prop]);
        }
      }
      return values;
    };

    var i = fuzzyset.gramSizeLower;
    for (i; i < fuzzyset.gramSizeUpper + 1; ++i) {
      fuzzyset.items[i] = [];
    }
    for (i = 0; i < arr.length; ++i) {
      fuzzyset.add(arr[i]);
    }

    return fuzzyset;
  };

  return FuzzySet;
})();

/**
 * Calculates the fuzzy match score between two strings.
 * @param {string} needle - The string to search for
 * @param {string} haystack - The string to search in
 * @returns {number} Match score from 0 to 100
 */
function fuzzyMatchScore(needle, haystack) {
  let fuzzySet = FuzzySet();
  fuzzySet.add(haystack);
  let result = fuzzySet.get(needle);
  if (!result) return 0;
  return result[0][0] * 100;
}

// ============================================================================
// EXPORTS (for use as a module if needed)
// ============================================================================

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    shouldExcludeSearchTerm,
    fuzzyMatchScore,
    SearchTermMatcher
  };
}

// ============================================================================
// RUN THE TESTS
// ============================================================================

runTests();
