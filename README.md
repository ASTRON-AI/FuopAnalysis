# MongoDB and MongoDB Compass Installation Guide

This guide provides step-by-step instructions on how to install MongoDB and MongoDB Compass on your system.

## Table of Contents

1. [MongoDB Installation](#mongodb-installation)
    - [Prerequisites](#prerequisites)
    - [Installation](#installation)
    - [Verification](#verification)
2. [MongoDB Compass Installation](#mongodb-compass-installation)
    - [Installation](#installation-1)
    - [Verification](#verification-1)
3. [Connecting MongoDB Compass to MongoDB](#connecting-mongodb-compass-to-mongodb)

## MongoDB Installation

### Prerequisites

Before installing MongoDB, ensure you have the following:

- A supported operating system:
  - Windows 7 and later
  - macOS 10.10 and later
  - Linux (RHEL, CentOS, Ubuntu, Debian, SUSE)
- Administrative privileges on your system
- An internet connection

### Installation

#### Windows

1. **Download the MongoDB Installer**
   - Go to the [MongoDB Download Center](https://www.mongodb.com/try/download/community) and download the Windows installer.

2. **Run the Installer**
   - Double-click the downloaded `.msi` file and follow the setup instructions.
   - Choose "Complete" for a full installation.
   - Ensure "Install MongoDB as a Service" is checked.

3. **Add MongoDB to the System Path**
   - Open the Control Panel.
   - Go to System and Security > System > Advanced system settings.
   - Click on "Environment Variables".
   - Under "System variables", select the `Path` variable and click "Edit".
   - Add the path to the MongoDB `bin` directory (e.g., `C:\Program Files\MongoDB\Server\4.4\bin`).

#### macOS

1. **Install Homebrew**
   - If you don't have Homebrew installed, open the Terminal and run:
     ```sh
     /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
     ```

2. **Install MongoDB**
   - Run the following command in the Terminal:
     ```sh
     brew tap mongodb/brew
     brew install mongodb-community
     ```

3. **Start MongoDB**
   - Run:
     ```sh
     brew services start mongodb/brew/mongodb-community
     ```

#### Linux

1. **Import the Public Key**
   - Run the following command to import the MongoDB public GPG key:
     ```sh
     wget -qO - https://www.mongodb.org/static/pgp/server-4.4.asc | sudo apt-key add -
     ```

2. **Create the List File**
   - Create the `/etc/apt/sources.list.d/mongodb-org-4.4.list` file for your version of Ubuntu:
     ```sh
     echo "deb [ arch=amd64,arm64 ] https://repo.mongodb.org/apt/ubuntu focal/mongodb-org/4.4 multiverse" | sudo tee /etc/apt/sources.list.d/mongodb-org-4.4.list
     ```

3. **Install MongoDB Packages**
   - Run:
     ```sh
     sudo apt-get update
     sudo apt-get install -y mongodb-org
     ```

4. **Start MongoDB**
   - Run:
     ```sh
     sudo systemctl start mongod
     ```

### Verification

1. Open a terminal or command prompt.
2. Run `mongo` to start the MongoDB shell.
3. If you see the MongoDB shell prompt (`>`), MongoDB is successfully installed and running.

## MongoDB Compass Installation

### Installation

1. **Download MongoDB Compass**
   - Go to the [MongoDB Compass Download Page](https://www.mongodb.com/try/download/compass) and download the installer for your operating system.

2. **Install MongoDB Compass**
   - Run the downloaded installer and follow the installation instructions.

### Verification

1. Open MongoDB Compass.
2. If the application opens successfully, MongoDB Compass is installed.




# Project Setup and Execution Guide

This guide provides step-by-step instructions on how to install a conda environment and run `main.py`.

## Table of Contents

1. [Conda Environment Installation](#conda-environment-installation)
    - [Prerequisites](#prerequisites)
    - [Creating the Conda Environment](#creating-the-conda-environment)
    - [Activating the Conda Environment](#activating-the-conda-environment)
2. [Running `main.py`](#running-mainpy)

## Conda Environment Installation

### Prerequisites

Before setting up the conda environment, ensure you have the following:

- Anaconda or Miniconda installed on your system. You can download and install Miniconda from [here](https://docs.conda.io/en/latest/miniconda.html) or Anaconda from [here](https://www.anaconda.com/products/distribution).

### Creating the Conda Environment

1. **Open a Terminal or Command Prompt**
2. **Create a conda Environment**
   ```sh
   conda create --name <YourEnvName> python=X.X 
3. **Install the Environment Packages**
   ```sh
   pip install requiresments.txt

## Running `main.py`

1. **Ensure the Conda Environment is Activated**

    ```sh
    conda activate myenv
    ```

2. **Navigate to the Project Directory**

    ```sh
    cd /path/to/your/project
    ```

3. **Run `main.py`**

    ```sh
    python main.py
    ```

   If `main.py` requires any command-line arguments, include them after `main.py`. For example:

    ```sh
    python main.py arg1 arg2
    ```

## Starting the Flask Web Application

1. **Ensure the Conda Environment is Activated**

    ```sh
    conda activate myenv
    ```

2. **Navigate to the Project Directory**

    ```sh
    cd /path/to/your/project
    ```

3. **Run `app.py`**

    ```sh
    python app.py
    ```

4. **Access the Web Application**

   Once the application is running, open your web browser and go to:

    ```
    http://localhost:5002
    ```

---
# write_todb.py

This script is designed to extract data from an Excel file and store it in a MongoDB database. The script specifically reads data from a specified sheet in the Excel file, processes it, and then inserts it into a MongoDB collection.
 ```sh
    python write_todb.py
 ```
