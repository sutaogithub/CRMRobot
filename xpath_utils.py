import configparser
import codecs


class XPathHelper:

    def __init__(self, file_path):
        self.path = file_path
        self.config = configparser.ConfigParser()
        self.config.read_file(open(file_path, encoding="UTF-8"))
        self.combobox_input = self.config.get("input", "combobox_input")
        self.editor_input = self.config.get("input", "editor_input")
        self.button_input = self.config.get("input", "button_input")

    def get_wifi_xpath(self, key):
        return self.config.get("wifi", key)

    def get_itv_xpath(self, key):
        return self.config.get("itv", key)

    def get_adsl_xpath(self, key):
        return self.config.get("adsl", key)

    def get_frame_xpath(self, key):
        return self.config.get("frame", key)

# xpath = XPathHelper("xpath.ini")
# print(xpath.get_wifi_xpath("wifi_location"))
