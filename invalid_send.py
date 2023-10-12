# I'm thinking about using this with the not_valid_emails list to send emails to the invalid emails with another program if they become too rampant    


#Using the pickle library to save a a list from one python program and then open that file in another program 

#Taking the orignal list and storing it using pickle
import pickle

my_list = [1, 2, 3, 4, 5]

with open('my_list.pkl', 'wb') as f:
    pickle.dump(my_list, f)

#loading the stored list into another program
import pickle

with open('my_list.pkl', 'rb') as f:
    my_list = pickle.load(f)

print(my_list)
