from score_calculator import Item


def start(girls, boys, attribution):
    def handleAffectation(girl):
        affected = False
        choice = 0
        girl_list = girls[girl]
        girl_list.sort(key=lambda x: x.rank, reverse=True)
        while affected == False:
            current_choice = girl_list[choice]
            current_choice_ranking = boys[current_choice.name]
            if attribution[current_choice.name] == None:
                attribution[current_choice.name] = girl
                affected = True
            elif attribution[current_choice.name] != None and attribution[current_choice.name] != girl:
                current_attributed_id = attribution[current_choice.name]
                current_ranked = current_choice_ranking[current_attributed_id]
                my_rank = current_choice_ranking[girl]
                if my_rank > current_ranked:
                    attribution[current_choice.name] = girl
                    affected = True
                    handleAffectation(current_attributed_id)
                elif my_rank <= current_ranked:
                    choice += 1



    for girl in girls:
        handleAffectation(girl)

    return attribution



