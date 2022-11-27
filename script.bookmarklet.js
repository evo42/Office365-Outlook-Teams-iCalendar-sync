
// nodejs script process input as `encodeURI` string

const fs = require('fs'),
        args = process.argv.slice(2)

//---debug---
//console.log(args)

let input = args[0] || './xyz.dist__bookmarklet-src.min.js',
    template = args[1] || './docs/index.htm',
    bookmarkletOutput = './docs/index.htm',
    encodedInput = null,
    templateData = null

if (input) {
    fs.readFile(input, 'utf8', function (err, data) {
      if (err) {
        return console.log(err)
      }

      //---debug---console.log(data);

      // URI Encode input data
      encodedInput = encodeURI(data)

      //---debug---console.log('*** encodedInput: ', encodedInput)

      try {
        fs.writeFileSync('./xyz.dist__app.encoded.js.txt', encodedInput);
        //---debug---
        console.log('\t - encodedInput file written successfully.')

        fs.readFile(template, 'utf8', function (err, content) {
          if (err) {
            return console.log(err)
          }

          //---debug---console.log(data);

          // URI Encode input data
          //
          // templateData = content.replace('alert(document.cookie);', encodedInput).replace('Get Biscuit.', 'iCal sync')
          templateData = content.replace('/<a href="javascript:(.*)\(\);">/g', '<a href="javascript: (() => { ' + encodedInput + ' })();')

          //---debug---console.log('*** templateData: ', templateData)

          try {
            fs.writeFileSync(bookmarkletOutput, templateData);
            //---debug---
            console.log('\t - templateData file written successfully.')
          } catch (err) {
            console.error(err)
          }
        })
      } catch (err) {
        console.error(err)
      }
    })
}