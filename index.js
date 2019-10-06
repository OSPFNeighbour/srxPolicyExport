const LineByLineReader = require('line-by-line')
const fs = require('fs');
const Excel = require('exceljs');

const showSecurityDisSetFile = 'security.sets.srx'
const showSecurityPoliciesHitCountFile = 'policy.hits.srx'


var lr = new LineByLineReader(showSecurityDisSetFile);


var zones = {}
var policies = {}


lr.on('error', function(err) {
  // 'err' contains error object
});

lr.on('line', function(line) {
  // pause emitting of lines...
  lr.pause();
  let lineParts = line.split(' ')
  //console.log(lineParts);
  // set security <foobar>
  switch (lineParts[2]) {
    case "zones": //for everything in the zone
      let zoneName = lineParts[4]
      //  console.log('a zone,', zoneName)
      if (!zones.hasOwnProperty(zoneName)) {
        zones[zoneName] = {}
        zones[zoneName]['address-book'] = []
      }
      // set security zone security-zone <foobar> address-book
      switch (lineParts[5]) {
        case "address-book":
          //console.log("address-book",line)
          if (lineParts[6] == "address") {
            let addressName = lineParts[7]
            let addressAddress = lineParts[8]
            let tmpBookEntry = {}
            tmpBookEntry.name = addressName
            tmpBookEntry.addresses = [addressAddress]
            zones[zoneName]['address-book'].push(tmpBookEntry)
          } else if (lineParts[6] == "address-set") {
            let addressSetName = lineParts[7]
            let addressSetEntryName = lineParts[9]
            // walk the current address books and resolve address-set entry
            // assumes addresses are before address sets
            let resolvedAddress = ''
            zones[zoneName]['address-book'].forEach(function(addr) {
              //  console.log(lineParts[9])
              if (addr.name == addressSetEntryName) {
                resolvedAddress = addr.addresses[0]
              }
            })
            if (resolvedAddress == '') {
              console.warn("An address-set address was not resolved", lineParts[7]);
            }
            // walk the current address entrys and add to the existing
            // or make a new zone
            let found = false
            zones[zoneName]['address-book'].forEach(function(addr) {
              if (addr.name == addressSetName) {
                found = true
                addr.addresses.push(resolvedAddress)
              }
            })
            if (!found) {
              let tmpBookEntry = {}
              tmpBookEntry.name = addressSetName
              tmpBookEntry.addresses = [resolvedAddress]
              zones[zoneName]['address-book'].push(tmpBookEntry)
            }
          }
          break
      }
      break
    case "policies": //for everything in the policices
      let fromZone = lineParts[4]
      let toZone = lineParts[6]
      let fromZoneToZone = fromZone + '->' + toZone
      let policyName = lineParts[8]

      if (!policies.hasOwnProperty(fromZoneToZone)) {
        policies[fromZoneToZone] = {}
      }

      if (!policies[fromZoneToZone].hasOwnProperty(policyName)) {
        policies[fromZoneToZone][policyName] = {}
        policies[fromZoneToZone][policyName]['from-zone'] = fromZone //because we will need this later
        policies[fromZoneToZone][policyName]['to-zone'] = toZone //because we will need this later
        policies[fromZoneToZone][policyName]['match'] = {}
        policies[fromZoneToZone][policyName]['match']['source-address'] = []
        policies[fromZoneToZone][policyName]['match']['destination-address'] = []
        policies[fromZoneToZone][policyName]['match']['application'] = []
        policies[fromZoneToZone][policyName]['then'] = []
      }

      if (lineParts[9] == "description") {
        //stupid spaces in descsriptions. TODO: honor double quotes in split
        policies[fromZoneToZone][policyName]['description'] = line.substring(line.indexOf('"') + 1, line.lastIndexOf('"'));
      } else if (lineParts[9] == "match") {
        //this could be done by using the [10] position for the key of the match like this...
        //but then it might go out of bounds if theres weird stuff
        //policies[fromZoneToZone][policyName]['match'][lineParts[10]].push(lineParts[11])
        switch (lineParts[10]) {
          case "source-address":
            policies[fromZoneToZone][policyName]['match']['source-address'].push(lineParts[11])
            break
          case "destination-address":
            policies[fromZoneToZone][policyName]['match']['destination-address'].push(lineParts[11])
            break
          case "application":
            policies[fromZoneToZone][policyName]['match']['application'].push(lineParts[11])
            break
        }
      } else if (lineParts[9] == "then") {
        if (lineParts[10] == "log") {
          policies[fromZoneToZone][policyName]['then'].push(line.substring(line.indexOf('log'),))
        } else {
          policies[fromZoneToZone][policyName]['then'].push(lineParts[10]) //permit or something single worded
        }
      } else {
        console.warn("NFI what this policy action is because its not (desc|match|then)", line)
      }
      break
  }
  lr.resume();
});

lr.on('end', function() {
  console.log('Done Reading Policy File')


  var hitslr = new LineByLineReader(showSecurityPoliciesHitCountFile);

  hitslr.on('error', function(err) {
    // 'err' contains error object
  });

  hitslr.on('line', function(line) {
    hitslr.pause();
    // pause emitting of lines...
    let cleanLine = line.replace(/\s\s+/g, ' ').split(' ')
    if (cleanLine.length == 7) { //a valid line
      if (policies.hasOwnProperty(cleanLine[2] + '->' + cleanLine[3])) {
        policies[cleanLine[2] + '->' + cleanLine[3]][cleanLine[4]]['hits'] = cleanLine[5]
      }
    }
    hitslr.resume();
  })

  hitslr.on('end', function() {
    console.log('Done Reading Policy Hits File')

    // Because policy is before zone we now have to walk everything again to resolve the policy addressAddress. fuck you JunOS
    Object.keys(policies).forEach(function(fromZoneToZone) {
      Object.keys(policies[fromZoneToZone]).forEach(function(policyName) {
        // it just assumes these match things exist. pretty sure JunOS forces that so thats cool.
        policies[fromZoneToZone][policyName]['match']['source-address-resolved'] = []
        policies[fromZoneToZone][policyName]['match']['destination-address-resolved'] = []
        policies[fromZoneToZone][policyName]['match']['source-address'].forEach(function(policySourceAddr) {
          if (policySourceAddr == "any") {
            policies[fromZoneToZone][policyName]['match']['source-address-resolved'].push('any')
          } else {
            zones[policies[fromZoneToZone][policyName]['from-zone']]['address-book'].forEach(function(bookAddr) {
              if (bookAddr.name == policySourceAddr) {
                policies[fromZoneToZone][policyName]['match']['source-address-resolved'] = policies[fromZoneToZone][policyName]['match']['source-address-resolved'].concat(bookAddr.addresses) //arrays for days
              }
            })
          }
        })
        policies[fromZoneToZone][policyName]['match']['destination-address'].forEach(function(policyDestAddr) {
          if (policyDestAddr == "any") {
            policies[fromZoneToZone][policyName]['match']['destination-address-resolved'].push('any')
          } else {
            zones[policies[fromZoneToZone][policyName]['to-zone']]['address-book'].forEach(function(bookAddr) {
              if (bookAddr.name == policyDestAddr) {
                policies[fromZoneToZone][policyName]['match']['destination-address-resolved'] = policies[fromZoneToZone][policyName]['match']['destination-address-resolved'].concat(bookAddr.addresses) //arrays for days
              }
            })
          }
        })
      });
      //console.log(JSON.stringify(policies[fromZoneToZone], null, 2))
    });
    // All lines are read, file is closed now.
    //console.log(zones)
    var outputworkbook = new Excel.Workbook();

    //setup common collumn headers
    var defaultSheetHeaders = [{
        header: 'policy-name',
        key: 'policy-name',
        width: 35
      },
      {
        header: 'policy-desc',
        key: 'policy-desc',
        width: 100
      },
      {
        header: 'policy-hits',
        key: 'policy-hits',
        width: 20
      },
      {
        header: 'from-zone',
        key: 'from-zone',
        width: 15
      },
      {
        header: 'to-zone',
        key: 'to-zone',
        width: 15
      },
      {
        header: 'source-address',
        key: 'source-address',
        width: 50,
        outlineLevel: 1
      },
      {
        header: 'source-address-resolved',
        key: 'source-address-resolved',
        width: 50,
        outlineLevel: 1
      },
      {
        header: 'destination-address',
        key: 'destination-address',
        width: 50,
        outlineLevel: 1
      },
      {
        header: 'destination-address-resolved',
        key: 'destination-address-resolved',
        width: 50,
        outlineLevel: 1
      },
      {
        header: 'application',
        key: 'application',
        width: 50,
        outlineLevel: 1
      },
      {
        header: 'then',
        key: 'then',
        width: 50,
        outlineLevel: 1
      },
    ];

    //First sheet has all the rules, following sheets are zone to zone


    var outputSheet = outputworkbook.addWorksheet("All Rules"); //one sheet, walk all policies

    outputSheet.columns = defaultSheetHeaders

    let rowCount = 0 //track row number so we can alternate bg colors
    Object.keys(policies).forEach(function(fromZoneToZone) {
      Object.keys(policies[fromZoneToZone]).forEach(function(policyName) {
        var rowValues = []
        rowValues[1] = policyName
        rowValues[2] = policies[fromZoneToZone][policyName]['description']
        rowValues[3] = policies[fromZoneToZone][policyName]['hits']
        rowValues[4] = policies[fromZoneToZone][policyName]['from-zone']
        rowValues[5] = policies[fromZoneToZone][policyName]['to-zone']
        rowValues[6] = policies[fromZoneToZone][policyName]['match']['source-address'].join('\n')
        rowValues[7] = policies[fromZoneToZone][policyName]['match']['source-address-resolved'].join('\n')
        rowValues[8] = policies[fromZoneToZone][policyName]['match']['destination-address'].join('\n')
        rowValues[9] = policies[fromZoneToZone][policyName]['match']['destination-address-resolved'].join('\n')
        rowValues[10] = policies[fromZoneToZone][policyName]['match']['application'].join('\n')
        rowValues[11] = policies[fromZoneToZone][policyName]['then'].join('\n')
        let lastRow = outputSheet.addRow(rowValues);

        rowCount++
        if (rowCount % 2) {
          lastRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
              argb: 'F2F2F2'
            }
          }
        } else {
          lastRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
              argb: 'E1E1E1'
            }
          }
        }
      })
    }) //end zone to zone


    // This seems like a painful way to do this but google didnt give me a better answer
    // iterate over all current cells in this column and set text wrap
    for (let column = 5; column < 12; column++) {
      outputSheet.getColumn(column).eachCell(function(cell, rowNumber) {
        cell.alignment = {
          wrapText: true
        };
      });
    }



    Object.keys(policies).forEach(function(fromZoneToZone) { //each zone to zone make a new sheet
      var outputSheet = outputworkbook.addWorksheet(fromZoneToZone); // fromzone->tozone naming

      outputSheet.columns = defaultSheetHeaders

      let rowCount = 0

      Object.keys(policies[fromZoneToZone]).forEach(function(policyName) {
        var rowValues = []
        rowValues[1] = policyName
        rowValues[2] = policies[fromZoneToZone][policyName]['description']
        rowValues[3] = policies[fromZoneToZone][policyName]['hits']
        rowValues[4] = policies[fromZoneToZone][policyName]['from-zone']
        rowValues[5] = policies[fromZoneToZone][policyName]['to-zone']
        rowValues[6] = policies[fromZoneToZone][policyName]['match']['source-address'].join('\n')
        rowValues[7] = policies[fromZoneToZone][policyName]['match']['source-address-resolved'].join('\n')
        rowValues[8] = policies[fromZoneToZone][policyName]['match']['destination-address'].join('\n')
        rowValues[9] = policies[fromZoneToZone][policyName]['match']['destination-address-resolved'].join('\n')
        rowValues[10] = policies[fromZoneToZone][policyName]['match']['application'].join('\n')
        rowValues[11] = policies[fromZoneToZone][policyName]['then'].join('\n')
        let lastRow = outputSheet.addRow(rowValues);

        rowCount++
        if (rowCount % 2) {
          lastRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
              argb: 'F2F2F2'
            }
          }
        } else {
          lastRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
              argb: 'E1E1E1'
            }
          }
        }
      })
      // This seems like a painful way to do this but google didnt give me a better answer
      // iterate over all current cells in this sheet and set text wrap
      for (let column = 5; column < 12; column++) {
        outputSheet.getColumn(column).eachCell(function(cell, rowNumber) {
          cell.alignment = {
            wrapText: true
          };
        });
      }
    });

    var newFile = 'Policy Export #' + new Date() / 1000 + '.xlsx'
    outputworkbook.xlsx.writeFile(newFile)
      .then(function() {
        console.log("XLS write done")
      });


  });
});
