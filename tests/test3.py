class SingleInstanceMetaClass(type):
    def __init__(self, name, bases, dic):
        self.__single_instance = None
        super().__init__(name, bases, dic)
 
    def __call__(cls, *args, **kwargs):
        if cls.__single_instance:
            return cls.__single_instance
        single_obj = cls.__new__(cls)
        single_obj.__init__(*args, **kwargs)
        cls.__single_instance = single_obj
        return single_obj
 
 
class Setting(metaclass=SingleInstanceMetaClass):
    def __init__(self):
        self.db = 'MySQL'
        self.port = 3306
 
 
bar1 = Setting()
bar2 = Setting()
 
print(bar1 is bar2)
print(bar1.db, bar1.port)
bar1.db = 'ORACLE'
print(bar2.db, bar2.port)