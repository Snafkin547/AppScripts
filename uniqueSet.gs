function getSet(sourceData){
  const uniqueSet = new Set();

  for (const row of sourceData) {
    const value = row[0];
    if (value !== "") { 
      uniqueSet.add(value);
    }
  }
  return uniqueSet
}