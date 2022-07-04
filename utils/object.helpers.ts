function deleteUndefinedProps(o: any) : any{
  Object.keys(o).forEach(function (key) {
    if(typeof o[key] === 'undefined'){
      delete o[key];
    }
  });
  return o; // mutated object
}

export {
  deleteUndefinedProps,
}
