# Facial Expression Recognition

## Overview

This Python project utilizes a Convolutional Neural Network (CNN) to first train and test a model with a provided dataset, and then use it to capture and recognize facial expressions in real-time through the device's camera.

The project uses TensorFlow for model construction and training, and OpenCV for image capturing and recognition.

## Usage

## Usage

1. **Create a Folder and Set Up a Virtual Environment:**
    - Create a new folder for the project and navigate into it.
    - Open a terminal or command prompt window.

    ```bash
    python -m venv venv
    source venv/bin/activate  # For Linux/Mac
    # or
    .\venv\Scripts\activate  # For Windows
    ```

2. **Download Kaggle Dataset:**
    - Download the "FER-2013" dataset from Kaggle using the following [link](https://www.kaggle.com/datasets/msambare/fer2013) and save it inside your project folder.

3. **Install Required Libraries:**
    - In the terminal or command prompt, run the following command:

    ```bash
    pip install numpy opencv-python pillow scikit-learn tensorflow
    ```

4. **Choose Model Usage Option:**
    - You can either train the model using the provided dataset or use a pre-trained model.

5. **Train the Model (Optional):**
    - Download the `main.py` file and place it in your project folder.
    - Open a terminal or command prompt and navigate to your project folder.

    ```bash
    python main.py
    ```

6. **Use the Model:**
    - Download the `use.py` and `haarcascade.xml` files and place them in your project folder.
    - If you trained the model, use the trained model file.
    - If you didn't train the model, download the `model.h5` file.

7. **Run the Model Usage Script:**
    - Open a terminal or command prompt and navigate to your project folder.

    ```bash
    python use.py
    ```

    - Press 'q' to exit the script when done.
