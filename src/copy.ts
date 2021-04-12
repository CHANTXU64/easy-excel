export function copyObject<T extends Object> (source_object: T): T {
  let key: keyof T;
  let target_object: T = Object.assign({}, source_object);
  for (key in source_object) {
    if (Object.prototype.hasOwnProperty.call(source_object, key)) {
      const source_element = source_object[key];
      if (source_element instanceof Array) {
        target_object[key] = copyArray(source_element);
      } else if (source_element instanceof Object) {
        target_object[key] = copyObject(source_element);
      }
    }
  }
  return target_object;
}

export function copyArray<T extends Array<any>> (source_arr: T): T {
  let length = source_arr.length;
  let target_arr = Object.assign([], source_arr);
  for (let i = 0; i < length; ++i) {
    let source_element = source_arr[i];
    if (source_element instanceof Array) {
      target_arr[i] = copyArray(source_element);
    } else if (source_element instanceof Object) {
      target_arr[i] = copyObject(source_element);
    } else {
      target_arr[i] = source_arr[i];
    }
  }
  return target_arr;
}

