// Branch: feature/item-mapping - Map item IDs to names using API and static overrides
/**
 * Item Mapping for Grand Exchange Farming Dashboard
 *
 * Contains static item name mappings and logic to fetch item mappings from the RuneScape API.
 */

// Static mapping of item names (overrides/corrections)
const itemNameMapping = {
  // Example: 'grimy ranarr weed': 'Ranarr seed',
  // Add more as needed
};

/**
 * Fetches item mappings from the RuneScape API.
 * Returns an object mapping itemId to itemName.
 */
function fetchItemMappings() {
  const url = 'https://prices.runescape.wiki/api/v1/osrs/mapping';
  try {
    const response = UrlFetchApp.fetch(url);
    const items = JSON.parse(response.getContentText());
    const itemMappings = {};
    items.forEach(item => {
      itemMappings[item.id] = item.name;
    });
    return itemMappings;
  } catch (error) {
    Logger.log('Error fetching item mappings: ' + error.toString());
    return {};
  }
}

/**
 * Gets the display name for an item, using API mapping and static overrides.
 */
function getDisplayName(itemId, apiMappings) {
  let name = apiMappings[itemId] || 'Unknown Item';
  const normalized = name.toLowerCase().trim();
  if (itemNameMapping[normalized]) {
    name = itemNameMapping[normalized];
  }
  return name;
} 