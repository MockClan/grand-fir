
const UnitTestingTools = {
  /**
    * Performs an equality test otherwise throws an exception indicating the lack of equality.
    * @param {string} name The name of the value being tested.
    * @param {string} actual The actual value.
    * @param {string} expected The expected value.
    */
  areEqual: function(name, actual, expected) {
    if (actual != expected) {
      throw new Error(`Invalid ${name} value - expected '${expected}', actual '${actual}'`);
    }
  }
};
