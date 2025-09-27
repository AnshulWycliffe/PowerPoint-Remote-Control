# PowerPoint Remote Control 🖥️

Control your PowerPoint presentations wirelessly from any device on your local network.

### WATCH DEMO VIDEO
[![▶ Watch the demo](https://img.youtube.com/vi/ZkbFl5UCtuA/0.jpg)](https://www.youtube.com/watch?v=ZkbFl5UCtuA)


## ✨ Features

*   📱 **Wireless Control**: Advance or go back in your presentation from any device (phone, tablet, laptop) connected to the same local network.
*   🌐 **Web-Based Interface & QR Functionality**: Simple and intuitive web interface, no special apps required on the remote device through QR or mannual.
*   🚀 **Easy Setup**: Get up and running quickly with minimal configuration.
*   💻 **Cross-Platform**: Works with any operating system that can run Python and access a web browser.
*   ⚡ **Lightweight**: Built with Flask, ensuring a small footprint and fast response times.


## 🛠️ Installation Guide

Follow these steps to set up and run the PowerPoint Remote Control on your system.

### Prerequisites

*   Python 3.12 installed on your machine.
*   `pip` (Python package installer).

### Step-by-Step Installation

1.  **Clone the Repository**:
    First, clone the project repository to your local machine using Git:

    ```bash
    git clone https://github.com/AnshulWycliffe/PowerPoint-Remote-Control.git
    cd PowerPoint-Remote-Control
    ```

2.  **Create a Virtual Environment (Recommended)**:
    It's good practice to create a virtual environment to manage project dependencies.

    ```bash
    python -m venv venv
    ```

3.  **Activate the Virtual Environment**:
    *   **On Windows**:

        ```bash
        .\venv\Scripts\activate
        ```

    *   **On macOS/Linux**:

        ```bash
        source venv/bin/activate
        ```

4.  **Install Dependencies**:
    Install all required Python packages using `pip` and the `requirements.txt` file:

    ```bash
    pip install -r requirements.txt
    ```

5.  **Run the Application**:
    Start the Flask application. This will launch a local web server.

    ```bash
    python app.py
    ```

    You should see output similar to this, indicating the server is running:

    ```
     * Serving Flask app 'app'
     * Debug mode: off
    WARNING: This is a development server. Do not use it in a production deployment. Use a production WSGI server instead.
     * Running on http://0.0.0.0:5000
    Press CTRL+C to quit
    ```


## 🚀 Usage Examples

Once the application is running, you can access the remote control interface from any device on your local network.

1.  **Open PowerPoint**: Ensure your PowerPoint presentation is open and in presentation mode on the host computer where the `app.py` is running.

2.  **Access the Web Interface**:
    On your remote device (e.g., smartphone, tablet), open a web browser and navigate to the IP address of the host computer, followed by port `5000`.

    For example, if your host computer's IP address is `192.168.1.100`, you would go to:

    ```
    http://192.168.1.100:5000
    ```

    You might also be able to use `http://localhost:5000` on the host machine itself.

3.  **QR Code Access**:
    When the server starts, a medium-sized QR code will be displayed in the terminal.

    Scan this QR code with your phone or tablet camera to open the remote interface directly—no typing needed.

4.  **Control Your Presentation**:
    The web interface will display simple buttons for navigation.
    *   Click `Next Slide` to advance to the next slide.
    *   Click `Previous Slide` to go back to the previous slide.

    ![Usage Screenshot Placeholder](/demo/image_1.png)

    ![Usage Screenshot Placeholder](/demo/image_2.png)


## 🗺️ Project Roadmap

This project is continuously evolving. Here are some planned features and improvements:

*   **Version 1.1.0**
    *   ⏱️ **Presentation Timer**: Display a timer on the remote interface.
    *   📝 **Speaker Notes Display**: Show speaker notes on the remote device.
    *   🔢 **Specific Slide Navigation**: Ability to jump to a specific slide number.
    *   🎨 **Customizable Themes**: Options for different UI themes on the remote control.
*   **Future Enhancements**
    *   🔄 **Real-time Slide Preview**: Display a small thumbnail of the current slide on the remote.
    *   📊 **Integration with other presentation software**: Extend support beyond PowerPoint (e.g., Google Slides, Keynote).
    *   🔒 **Basic Authentication**: Add optional password protection for the remote interface.


## 🤝 Contribution Guidelines

We welcome contributions to make this project even better! Please follow these guidelines:

1.  **Fork the Repository**: Start by forking the `PowerPoint-Remote-Control` repository to your GitHub account.
2.  **Create a New Branch**:
    *   For new features, create a `feature/your-feature-name` branch.
    *   For bug fixes, create a `bugfix/issue-description` branch.
3.  **Code Style**: Adhere to [PEP 8](https://www.python.org/dev/peps/pep-0008/) for Python code.
4.  **Commit Messages**: Write clear and concise commit messages.
    *   Use the present tense ("Add feature" instead of "Added feature").
    *   Start with an emoji related to the commit (e.g., ✨ for new features, 🐛 for bug fixes).
5.  **Testing**: Before submitting a Pull Request, ensure your changes work as expected and don't introduce new issues. Test thoroughly on your local machine.
6.  **Open a Pull Request**:
    *   Submit a Pull Request from your branch to the `main` branch of this repository.
    *   Provide a detailed description of your changes and why they are necessary.
    *   Reference any related issues.










