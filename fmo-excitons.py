import argparse


ham_help = """
The Hamiltonian must be specified in a CSV file with N rows and N columns where N is the number of pigments in the system.
"""

dip_help = """
The dipole moments must be specified in a CSV file with N rows and 7 columns. Each row corresponds to one pigment.
The first column is the magnitude of the dipole moment. Columns 2-4 indicate the x-, y-, and z-components respectively
of the transition dipole moment. Columns 5-7 indicate the x-, y-, and z-coordinates respectively of each pigment.
"""

config_help = """
The config file specifies parameters for the simulation. Details about the configuration file are shown at the bottom
of the help page.
"""

epilog = """
Configuration File
==================

The config file is a TOML file with a single section as shown below:

# config.toml
[parameters]
pigments = 7
bandwidth = 100
spectrum_start = 11600
spectrum_step = 1
spectrum_stop = 13100

All parameters are required. The [parameters] heading is also required. The simulation will be performed within the
interval [spectrum_start, spectrum_stop] with a step size of spectrum_step.
"""

def main():
    pass


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Compute absorption and CD spectra for FMO",
        epilog=epilog,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("-H", "--hamiltonian", type=argparse.FileType("r"), help=ham_help, metavar="FILE")
    parser.add_argument("-D", "--dipoles", type=argparse.FileType("r"), help=dip_help, metavar="FILE")
    parser.add_argument("-c", "--config", type=argparse.FileType("r"), help=config_help, metavar="FILE")
    args = parser.parse_args()
    main()
