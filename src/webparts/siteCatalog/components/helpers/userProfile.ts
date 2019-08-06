const getUserProfileProperty = (properties: Array<any>, key: string):string => {

    for (let x = 0; x < properties.length; x++) {
      if (properties[x].Key === key) {
        return properties[x].Value;
      }
    }
  
    return null;
  };
  
  export { getUserProfileProperty };