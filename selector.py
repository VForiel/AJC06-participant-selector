import pandas as pd
import os
import yaml
import numpy as np

#==============================================================================
# Person class
#==============================================================================

class Person:
    def __init__(self, id, first_name, family_name, email, phone, gender, ed, have_car, places, already_attended):
        self.id = id
        self.first_name = first_name
        self.family_name = family_name
        self.email = email
        self.phone = phone
        self.gender = gender
        self.ed = ed
        self.have_car = have_car
        self.places = places
        self.already_attended = already_attended

    @property
    def id(self):
        return self._id
    @id.setter
    def id(self, value):
        self._id = int(value)

    @property
    def first_name(self):
        return self._first_name[0].upper() + self._first_name[1:]
    @first_name.setter
    def first_name(self, value):
        self._first_name = value.replace(" ", "").lower()

    @property
    def family_name(self):
        return self._family_name[0].upper() + self._family_name[1:]
    @family_name.setter
    def family_name(self, value):
        self._family_name = value.replace(" ", "").lower()

    @property
    def email(self):
        return self._email
    @email.setter
    def email(self, value):
        self._email = value.replace(" ", "").lower()

    @property
    def phone(self):
        return self.phone
    @phone.setter
    def phone(self, value):
        self._phone = value.replace(" ", "").replace(".", "")

    @property
    def gender(self):
        return self._gender
    @gender.setter
    def gender(self, value):
        self._gender = value

    @property
    def ed(self):
        return self._ed
    @ed.setter
    def ed(self, value):
        self._ed = value if value in ["ED DESPEG", "ED SFA", "ED SHAL", "ED SMH", "ED STIC", "ED SVS"] else "Other"

    @property
    def have_car(self):
        return self._have_car
    @have_car.setter
    def have_car(self, value):
        self._have_car = True if value == "Yes" or "If necessary" else False

    @property
    def places(self):
        return self._places
    @places.setter
    def places(self, value):
        try:
            self._places = int(value)
        except ValueError:
            self._places = 0

    @property
    def already_attended(self):
        return self._already_attended
    @already_attended.setter
    def already_attended(self, value):
        self._already_attended = True if value == "Yes" else False

    @property
    def full_name(self):
        return self.first_name + " " + self.family_name

    def __str__(self):
        return f"{self.full_name} ({self.email})"
    
    def __repr__(self):
        return f"{self.full_name} ({self.email})"


#==============================================================================
# Parse the .xlsx file
#==============================================================================

# Find the .xlsx file
matches = []
for file in os.listdir("."):
    if file.endswith(".xlsx") and not file.startswith("~$"):
        matches.append(file)

if len(matches) == 0:
    print("No .xlsx files found.")
    exit(1)

if len(matches) > 1:
    print("Multiple .xlsx files found. Please remove all unwanted files.")
    exit(1)

# Load pre-registration data
preregistered = pd.read_excel(matches[0])
pool = [
    Person(
        id = i+2,
        first_name = preregistered["First name"][i],
        family_name = preregistered["Family name"][i],
        email = preregistered["Email Address"][i],
        phone = preregistered["Phone number"][i],
        gender = preregistered["Gender"][i],
        ed = preregistered["Doctoral school"][i],
        have_car = preregistered["Do you plan to take your car to come? (winter equipment might be required)"][i],
        places = preregistered["How many people can you carry in your car? (besides you)"][i],
        already_attended = preregistered["Have you attended a previous winter camp? "][i]
    )
    for i in range(len(preregistered["First name"]))
]

#==============================================================================
# Manage config
#==============================================================================

# Load the config data
config = yaml.safe_load(open("config.yaml"))

places = config["places"]

organizers = [pool[i-2] for i in config["organizers"]]
conflicts = []
for conflict in config["conflicts"]:
    conflicts.append([pool[i-2] for i in conflict])
groups = []
for group in config["groups"]:
    groups.append([pool[i-2] for i in group])

already_selected = [pool[i-2] for i in config["registered"]] if "registered" in config else []

#==============================================================================
# Registering functions
#==============================================================================

registered = []
rejected = []

def remove(ppl):
    try:
        pool.remove(ppl)
    except ValueError:
        pass

def remove_conflict(ppl):
    for conflict in conflicts:
        if ppl in conflict:
            for c in conflict:
                if c in pool and c is not ppl:
                    print("- Removed " + c.full_name + " from the pool due to conflict with " + ppl.full_name)
                    rejected.append(c)
                    pool.remove(c)

def register_group(ppl):
    for group in groups:
        if ppl in group:
            for g in group:
                if g in pool and g is not ppl:
                    print("   -> Registering " + g.full_name + " because he/she is grouped with " + ppl.full_name)
                    register(g)

def register(ppl):
    print("+ registered " + ppl.full_name)
    remove(ppl)
    remove_conflict(ppl)
    if ppl not in registered:
        registered.append(ppl)
    register_group(ppl)

#==============================================================================
# Register poeple
#==============================================================================

print("\nAdding organizers:")
for ppl in organizers:
    register(ppl)

print("\nAdding already selected people:")
for ppl in already_selected:
    register(ppl)

sits = 0

# Compute the probability of a people to be selected
def compute_prob(ppl):
    prob = 1

    # If there is not enough sits, favor people with a car
    if sits < len(registered):
        if not ppl.have_car:
            prob *= 1e-3 # Low probability for people without car
    
    prob *= max(1, ppl.places) # Increased probability for people who can bring others
    
    # ED equality
    same = 0
    other = 0
    if ppl.ed != "Other":
        for p in registered:
            if p.ed == ppl.ed:
                same += 1
            else:
                other += 1
        same = max(same, 1e-2)
        other = max(other, 1e-2)
        prob *= other / same

    # Compute gender ratio
    if ppl.gender != "Non-binary":
        same = 0
        other = 0
        for p in registered:
            if p.gender == ppl.gender:
                same += 1
            else:
                other += 1
        same = max(same, 1e-2)
        other = max(other, 1e-2)    
        prob *= other / same

    # Reduce probability for people who already attended
    if ppl.already_attended:
        prob /= 3

    return prob

print("\nAdding participants:")
while len(registered) < 35:
    ids = [i for i in range(len(pool))]
    prob = [compute_prob(pool[i]) for i in ids]

    # Randomly select a people based on the probability
    selected_id = np.random.choice(ids, p=prob/np.sum(prob))
    register(pool[selected_id])
        
# Print the final list
print(f"\nregistered: ({len(registered)})")
for ppl in registered:
    print(" - ", ppl.full_name)

# Print the remaining people
print(f"\nRemaining people: ({len(pool) + len(rejected)})")
for ppl in pool:
    print(" - ", ppl.full_name)
for ppl in rejected:
    print(" - (Rejected)", ppl.full_name)

# Print gender ratio
print("\nGender ratio:")
males = 0
females = 0
for ppl in registered:
    if ppl.gender == "Male":
        males += 1
    elif ppl.gender == "Female":
        females += 1
print(" - Males:", round(males / len(registered) * 100,1), "%")
print(" - Females:", round(females / len(registered) * 100,1), "%")
print(" - Non-binary:", round((len(registered) - males - females) / len(registered) * 100,1), "%")

# Print ED ratio
print("\nED ratio:")
EDs = {}
for ppl in registered:
    if ppl.ed not in EDs:
        EDs[ppl.ed] = 0
    EDs[ppl.ed] += 1
for ed in EDs:
    print(" -", ed, ":", round(EDs[ed] / len(registered) * 100,1), "%")

# Print sits
sits = 0
for ppl in registered:
    sits += ppl.places
print("\nSits:", sits)
print("Remaining sits:", sits - len(registered))

# Ask to keep the result

choice = input("\nDo you want to keep the result? (y/N)")
if choice.lower() in ["y","yes"]:
    with open("config.yaml", "w") as conf_file:
        config["registered"] = [ppl.id for ppl in registered]
        yaml.dump(config, conf_file)

    # Email list
    print("\nEmail list:")
    print(", ".join([p.email for p in registered]))
else:
    print("Result not saved.")