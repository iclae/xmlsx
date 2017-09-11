function setHide(store, cell) {
  store.style({
    cell: cell,
    style: 'hide',
  })
}

module.exports = {
  setHide: setHide,
}
