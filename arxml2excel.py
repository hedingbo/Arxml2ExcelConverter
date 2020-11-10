import cantools





if __name__ == "__main__":
    db = cantools.database.load_file(r"200806-hd2-EP33L_Simu1_ICC-Test.arxml")
    print(db.version) 