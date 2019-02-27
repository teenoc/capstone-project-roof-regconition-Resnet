1.The classification models are built on fastai library, use the below link to install fastai library. 
https://github.com/fastai/fastai
once finished, use jupyter notebook to run the python 3 script for performing classification. 

2.Current models are both developed based on pretrianed ResNet101.
The code for retraining the ResNet101 and individual prediction 
are shown in result folder under each objective folder, 
given in the form of .ipynb and a printed pdf.

3.Training process in the script is done using GPU computing, 
modification to the code is required if CPU computing mode is chosen.

4.A word document related to choosing batch size and learning rate at different image size is included.

5.the test result are given in the test_result.xlsx file.

6.the updated weight resides under the 'model' folder, to do prediction, simply load type_predict.ipynb or roof_predict.ipynb.  

7.currently the model has not been fully automated yet, new address can be fed to 'new_address.xlsx', 
then run 'Image Extraction Script.py' on 64-bit IDLE to download Image to 'predict' folder,   
and run roof_predict.ipynb to do individual prediction or configure the code to do bulk prediciton.

8.Definition of 'Commercial' and 'Residential' building are given in the slides.

9.Definition of 'Membrane', 'Metal', 'Shingles', 'Tiles' are given in the slides.

