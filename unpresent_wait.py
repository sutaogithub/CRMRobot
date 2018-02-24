class unpresence_of_element_located(object):
    """An expectation for checking that an element has a particular css class.

    locator - used to find the element
    returns the WebElement once it has the particular css class
    """

    def __init__(self, locator):
        self.locator = locator

    def __call__(self, driver):
        try:
            driver.find_element(*self.locator)  # Finding the referenced element
            return False
        except Exception as e:
            return True
