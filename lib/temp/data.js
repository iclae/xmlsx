function getData() {
  return {
    sst: {
      $: {
        xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        count: '1',
        uniqueCount: '1',
      },
      si: [],
    },
  }
}

module.exports = {
  getData: getData,
}
