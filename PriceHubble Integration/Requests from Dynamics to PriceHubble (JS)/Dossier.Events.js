function sendUpdateRequestToPriceHubble(executionContext) {
    var formContext = executionContext.getFormContext();

    var countryCode = formContext.getAttribute("ph_countrycode");
    var dealType = formContext.getAttribute("ph_dealtype");
    var sourceDossierId = formContext.getAttribute("ph_sourcedossierid");
    var roomTourLink = formContext.getAttribute("ph_roomtourlink");
    var title = formContext.getAttribute("ph_title");
    var property = formContext.getAttribute("ph_propertyid");    

    var dossierId = formContext.data.entity.getId().replace('{', '').replace('}', '');
    var authToken = '74126eab0a9048d993bda4b1b55ae074';

    var updatedFields = { };

    if (countryCode && countryCode.getIsDirty()) {
        updatedFields["countryCode"] = countryCode.getValue();
    }
    else if (dealType && dealType.getIsDirty()) {
        updatedFields["dealType"] = dealType.getValue();
    }
    else if (sourceDossierId && sourceDossierId.getIsDirty()) {
        updatedFields["sourceDossierId"] = sourceDossierId.getValue();
    }
    else if (roomTourLink && roomTourLink.getIsDirty()) {
        updatedFields["roomTourLink"] = roomTourLink.getValue();
    }
    else if (title && title.getIsDirty()) {
        updatedFields["title"] = title.getValue();
    }
    else if (property && property.getIsDirty()) {
        var idProperty = property.getValue()[0].id;
        var record = Xrm.WebApi.retrieveRecord("ph_property", idProperty);
        updatedFields["property"] = {
            "location": {
                "address": {
                  "postCode": "8037",
                  "city": "Zurich",
                  "street": "Nordstrasse",
                  "houseNumber": "391"
                },
                "coordinates": {
                  "latitude": 47.3968601,
                  "longitude": 8.5153549
                }
              },
              "propertyType": {
                "code": "house",
                "subcode": "house_detached"
              },
              "buildingYear": record.ph_buildingyear,
              "balconyArea": record.ph_balconyarea,
              "hasLift":record.ph_haslift,
              "hasPool":record.ph_haspool,
              "hasSauna": record.ph_hassauna,
              "livingArea": record.ph_livingarea,
              "landArea": record.ph_landarea,
              "volume": record.ph_volume,
              "renovationYear": record.ph_renovationyear,
              "numberOfRooms": 3,
              "numberOfBathrooms": record.ph_numberofbathrooms,
              "numberOfIndoorParkingSpaces": record.ph_numberofindoorparkingspaces,
              "numberOfOutdoorParkingSpaces": record.ph_numberofoutdoorparkingspaces,
              "hasPool": true,
              "condition": {
                "bathrooms": "renovation_needed",
                "kitchen": "renovation_needed",
                "flooring": "well_maintained",
                "windows": "new_or_recently_renovated"
              },
              "quality": {
                "bathrooms": "simple",
                "kitchen": "normal",
                "flooring": "high_quality",
                "windows": "luxury"
              }
        };
    }

    // var apiUrl = 'https://api.pricehubble.com/api/v1/dossiers/' + dossierId;
    // var xhr = new XMLHttpRequest();
    // xhr.open('PATCH', apiUrl, true);
    // xhr.setRequestHeader('Content-Type', 'application/json');
    // xhr.setRequestHeader('Authorization', 'Bearer ' + authToken);
    // xhr.send(JSON.stringify(updatedFields));
}


