import os
import yaml
from sys import argv
import platform
from handlers.vulnerability import Vulnerability


def config_loader():
    config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.yaml")
    with open(config_file) as yaml_file:
        conf = yaml.load(yaml_file, Loader=yaml.FullLoader)
    return conf


def get_platform():
    return platform.system().lower()


def get_source_and_destination():
    try:
        source = argv[1]
        destination = argv[2]
        replace = False
        if "-n" in argv:
            replace = True
        return source, destination, replace
    except IndexError:
        print("Incorrect source or destination path.")
        exit(127)


def start():
    source, destination, replace = get_source_and_destination()
    config = config_loader()
    config['document']['replace'] = replace
    system_platform = get_platform()
    vulnerability = Vulnerability(config, source, destination, system_platform)
    vulnerability.process()


if __name__ == '__main__':
    start()
