export async function universalSearchQuery(
  searchText: string,
  searchableFields: string[],
) {
  if (!searchText) return {};
  const escapedSearch = searchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return {
    $or: searchableFields.map((field) => ({
      [field]: {
        $regex: escapedSearch,
        $options: 'i',
      },
    })),
  };
}
