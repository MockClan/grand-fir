
/**
 * USPS Unit Tests Namespace
 */
const UspsTests = {

  testAddressLookup: function() {
    var result = Usps.addressLookup('10375 Mount Wilson Place', 'Peyton', 'CO');
    console.log(JSON.stringify(result));
    var expected = new UspsAddress(null, '10375 MOUNT WILSON PL', 'PEYTON', 'CO', '80831', '4449');

    UnitTestingTools.areEqual('address1', result.address1, expected.address1);
    UnitTestingTools.areEqual('address2', result.address2, expected.address2);
    UnitTestingTools.areEqual('city', result.city, expected.city);
    UnitTestingTools.areEqual('state', result.state, expected.state);
    UnitTestingTools.areEqual('zip5', result.zip5, expected.zip5);
    UnitTestingTools.areEqual('zip4', result.zip4, expected.zip4);
  },

  testParseUspsZipCodeLookup: function() {
    // eslint-disable-next-line no-multi-str
    var xml = '<?xml version="1.0" encoding=\"UTF-8\"?>\n<ZipCodeLookupResponse>\
    <Address ID="0"><Address2>10375 MOUNT WILSON PL</Address2><City>PEYTON</City>\
    <State>CO</State><Zip5>80831</Zip5><Zip4>4449</Zip4></Address></ZipCodeLookupResponse>';

    var result = Usps._parseUspsZipCodeLookup(xml);
    var expected = new UspsAddress(null, '10375 MOUNT WILSON PL', 'PEYTON', 'CO', '80831', '4449');
    UnitTestingTools.areEqual('address1', result.address1, expected.address1);
    UnitTestingTools.areEqual('address2', result.address2, expected.address2);
    UnitTestingTools.areEqual('city', result.city, expected.city);
    UnitTestingTools.areEqual('state', result.state, expected.state);
    UnitTestingTools.areEqual('zip5', result.zip5, expected.zip5);
    UnitTestingTools.areEqual('zip4', result.zip4, expected.zip4);
  }
};

/**
 * Entry point for running the USPS functions unit tests.
 */
function _runUspsTests() {
  UspsTests.testAddressLookup();
  UspsTests.testParseUspsZipCodeLookup();
}

