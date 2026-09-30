import networkx as nx
import matplotlib.pyplot as plt

class DAGCreator:

    def __init__(self,filename):
        self.filename = filename
        self.graph = nx.DiGraph()

    def graph_creation(self, required_courses):
        for course in required_courses:
            self.graph.add_node(course)

        with open(self.filename, "r") as file:
            for line in file:
                line = line.strip()
                
                if not line:
                    continue
                
                if "->" not in line:
                    continue

                prereq , course = line.split("->")

                prereq = prereq.strip()
                course = course.strip()
            
                if course in required_courses:
                    self.graph.add_edge(prereq,course)

    def show_graph(self):
        if self.graph.number_of_nodes() == 0:
            print("No courses were found for the prerequisite graph.")
            return

        pos = nx.spring_layout(self.graph, seed=42)

        nx.draw(
            self.graph,
            pos,
            with_labels=True,
            node_size=2500,
            node_color="lightblue",
            arrows=True,
            font_size=9
        )

        plt.title("Student Prerequisite Graph")
        plt.show()

    def get_graph(self):
        return self.graph

    def get_prereq(self, course):
        return list(self.graph.predecessors(course))

    def get_courses(self):
        return list(self.graph.nodes)
       
