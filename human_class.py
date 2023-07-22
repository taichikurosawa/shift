member=[]

class human():#学生に関するクラス
    def __init__(self,name,species,grade,gender,club,setup,first,second):#情報で学生を定義
        self.name=name
        self.species=species#種別
        self.grade=grade#学年
        self.gender=gender
        self.club=club#部活 
        self.setup=setup#設営参加可否（oかx（半角英字のオーとエックス））
        self.first=first#1日目参加可否
        self.second=second#2日目参加可否
        self.number=1#枠数 

    #def remove(self,human):#リストから削除する
        #del member[member.index(human)]

    #def combine(self,invited):#紹介者に対して実行し、紹介者に紹介された学生の名前と枠を結合
        #self.name=self.name+"&"+invited.name
        #self.number+=invited.number
        #self.species=self.species+"&"+invited.species



class place():
    def __init__(self,name,number):
        self.name=name
        self.number=number#枠数
        self.member=[]#各場所のメンバー
    
    





