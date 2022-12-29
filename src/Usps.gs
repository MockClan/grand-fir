/**
 * USPS Address object definition
 */
class UspsAddress {
  /** Constructor
     * @param {string} address1 The subaddress for the location (suite, appt #, etc).
     * @param {string} address2 The street address for the location.
     * @param {string} city The city of the address.
     * @param {string} state The state code of the address.
     * @param {string} zip5 The Zip5 code of the address.
     * @param {string} zip4 The Zip4 code of the address.
     */
  constructor(address1, address2, city, state, zip5, zip4) {
    this.address1 = address1;
    this.address2 = address2;
    this.city = city;
    this.state = state;
    this.zip5 = zip5;
    this.zip4 = zip4;
  }

  /**
   * Returns a JSON serialized representation of the class.
   * @return {string} The JSON repesentation of the class.
   */
  toJson() {
    return JSON.stringify(this);
  }
}

/**
 * USPS Funcions Namespace
 */
const Usps = {
  /**
   * Function posts request to USPS address verification API and,
   * returns full address with Zip5 and Zip4 nodes
   * @param {string} address The street address to lookup.
   * @param {string} city The city of the address to lookup.
   * @param {state} state The state of the address to lookup.
   * @return {UspsAddress} An object representing the results of the lookup
   */
  addressLookup: function(address, city, state) {
    var apiURL = `http://production.shippingapis.com/ShippingAPI.dll?API=ZipCodeLookup&XML=\
    <ZipCodeLookupRequest USERID='${Secrets.Usps.UserId}'>\
    <Address ID='0'>\
    <FirmName></FirmName>\
    <Address1></Address1>\
    <Address2>${address}</Address2>\
    <City>${city}</City>\
    <State>${state}</State>\
    </Address></ZipCodeLookupRequest>`;

    var response = UrlFetchApp.fetch(encodeURI(apiURL)).getContentText();

    Logger.log(
        JSON.stringify({
          response: response
        })
    );

    return this._parseUspsZipCodeLookup(response);
  },

  _parseUspsZipCodeLookup: function(xml) {
    const addressElement = XmlService.parse(xml)
        .getRootElement()
        .getChild('Address');
    var address1 = addressElement.getChildText('Address1');
    var address2 = addressElement.getChildText('Address2');
    var city = addressElement.getChildText('City');
    var state = addressElement.getChildText('State');
    var zip5 = addressElement.getChildText('Zip5');
    var zip4 = addressElement.getChildText('Zip4');
    return new UspsAddress(address1, address2, city, state, zip5, zip4);
  }
};
