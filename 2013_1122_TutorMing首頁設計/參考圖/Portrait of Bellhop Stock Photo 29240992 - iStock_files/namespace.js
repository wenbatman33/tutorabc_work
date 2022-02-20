
namespacing={init:function(namespace){var spaces=[];namespace.split('.').each(function(space){var curSpace=window,i;spaces.push(space);for(i=0;i<spaces.length;i++){if(typeof curSpace[spaces[i]]==='undefined'){curSpace[spaces[i]]={};}
curSpace=curSpace[spaces[i]];}});}};